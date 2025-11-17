# backend_logic.py
import os
import re
import sys
import logging
import shutil
import tempfile
import traceback
import platform
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps
import pillow_heif
import qrcode

try:
    import win32print  # type: ignore
except ImportError:  # pragma: no cover - platform dependent
    win32print = None  # type: ignore[assignment]

try:
    import cups  # type: ignore
except ImportError:  # pragma: no cover - platform dependent
    cups = None  # type: ignore[assignment]

try:
    import barcode  # type: ignore
    from barcode.writer import ImageWriter  # type: ignore
    _BARCODE_IMPORT_ERROR: Optional[Exception] = None
except Exception as exc:  # pragma: no cover - handled at runtime for user feedback
    barcode = None  # type: ignore[assignment]
    ImageWriter = None  # type: ignore[assignment]
    _BARCODE_IMPORT_ERROR = exc


if hasattr(Image, "Resampling"):
    _RESAMPLE_NEAREST = Image.Resampling.NEAREST
else:  # pragma: no cover - Pillow < 9 compatibility
    _RESAMPLE_NEAREST = Image.NEAREST


if ImageWriter is not None:

    class _SafeImageWriter(ImageWriter):
        """Image writer that falls back to PIL's default font when unavailable.

        The python-barcode ``ImageWriter`` attempts to load a TrueType font for
        rendering the human readable text under the barcode. In minimal
        environments the bundled DejaVu fonts may be missing which would
        normally raise an :class:`OSError`. The application should keep
        rendering the barcode even in that scenario, therefore this helper
        wraps the font loading logic to gracefully fall back to Pillow's
        default bitmap font and records a user-friendly warning message.
        """

        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.font_fallback_used = False
            self.font_warning_message: Optional[str] = None

        def _load_font(self, font_path=None, size=10):  # type: ignore[override]
            candidates = []
            if font_path:
                candidates.append(font_path)

            for candidate in candidates:
                try:
                    return ImageFont.truetype(candidate, size=size)
                except OSError:
                    logging.debug("Unable to load barcode font '%s'", candidate)
                    continue

            try:
                font = ImageFont.truetype("DejaVuSans.ttf", size=size)
            except OSError:
                self.font_fallback_used = True
                self.font_warning_message = "Font not found, default font used."
                logging.warning(self.font_warning_message)
                font = ImageFont.load_default()
            return font

else:  # pragma: no cover - barcode library unavailable

    class _SafeImageWriter:  # type: ignore[too-few-public-methods]
        font_fallback_used = False
        font_warning_message: Optional[str] = None


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


def list_printers() -> List[str]:
    """Return available printers using platform-specific backends."""

    printers: List[str] = []
    system = platform.system()

    try:
        if system == "Windows":
            if win32print is None:
                return printers
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            for _, _, name, _ in win32print.EnumPrinters(flags):
                printers.append(name)
        else:
            if cups is None:
                return printers
            connection = cups.Connection()
            for name in connection.getPrinters().keys():
                printers.append(name)
    except Exception as exc:  # pragma: no cover - platform dependent
        logging.warning("Unable to list printers: %s", exc)

    return printers


def rinven_barcode_dependency_issue() -> Optional[str]:
    """Return a human friendly description if barcode generation is unavailable."""

    return _verify_barcode_dependencies()

def clean_file_path(file_path: str) -> str:
    """Cleans quotes and spaces from the file path string."""
    return file_path.strip().strip('"').strip("'")

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
    # Preserve the barcode's sharp edges when resizing for the preview label.
    code_image.thumbnail((max_code_w, max_code_h), _RESAMPLE_NEAREST)
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
    """Converts all HEIC and WEBP files in a folder to JPG."""
    log_callback("Starting HEIC/WEBP to JPG conversion...")
    try:
        files = [
            f
            for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f))
            and os.path.splitext(f)[1].lower() in {".heic", ".webp"}
        ]
        if not files:
            log_callback("No HEIC or WEBP files found."); return

        total_files = len(files)
        for i, f in enumerate(files):
            log_callback(f"Converting ({i+1}/{total_files}): {f}")
            src, dst = os.path.join(folder, f), f"{os.path.splitext(os.path.join(folder, f))[0]}.jpg"
            try:
                img = None
                ext = os.path.splitext(f)[1].lower()
                if ext == ".heic":
                    heif = pillow_heif.read_heif(src)
                    img = Image.frombytes(heif.mode, heif.size, heif.data, "raw")
                else:
                    with Image.open(src) as opened:
                        if opened.mode not in ("RGB", "L"):
                            img = opened.convert("RGB")
                        else:
                            img = opened.copy()

                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")

                img.save(dst, "JPEG")
            except Exception as e: log_callback(f"Error converting '{f}': {e}")
            finally:
                if img is not None:
                    try:
                        img.close()
                    except Exception:
                        pass

        log_callback("\n✅ Conversion complete.")
        completion_callback("Success", "HEIC/WEBP conversion is complete.")
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
    dependency_issue = _verify_barcode_dependencies()
    if dependency_issue is not None:
        return (f"Error generating barcode: {dependency_issue}", None)

    try:
        fallback_message: Optional[str] = None
        if output_type == "PNG":
            barcode_img, writer = _render_barcode_image(data, bc_format)
            fallback_message = getattr(writer, "font_warning_message", None)
            saved_fname = fname
            if not saved_fname.lower().endswith(".png"):
                saved_fname = f"{saved_fname}.png"
            barcode_img.save(saved_fname)
            log_message = f"✅ Barcode saved as '{saved_fname}'"
            detail_message = f"Barcode saved to:\n{os.path.abspath(saved_fname)}"
        else:  # Dymo
            barcode_pil_img, writer = _render_barcode_image(
                data,
                bc_format,
                writer_options={
                    "module_height": 18.0,
                    "quiet_zone": 6,
                },
            )
            fallback_message = getattr(writer, "font_warning_message", None)
            label_image = create_label_image(barcode_pil_img, dymo_size_info, bottom_text)
            label_image.save(fname)
            log_message = f"✅ Dymo Label saved as '{fname}'"
            detail_message = f"Dymo Label saved to:\n{os.path.abspath(fname)}"

        if fallback_message:
            log_message = f"⚠️ {fallback_message}\n{log_message}"

        return (log_message, detail_message)
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


def _append_barcode_text(barcode_img: Image.Image, text: str) -> Image.Image:
    """Append human readable text below the barcode image using fallback fonts."""

    text = (text or "").strip()
    if not text:
        return barcode_img

    if barcode_img.mode != "RGB":
        barcode_img = barcode_img.convert("RGB")

    width, height = barcode_img.size
    padding = max(4, width // 40)
    font_size = max(10, min(36, width // 10))
    font = _load_font(
        [
            "DejaVuSans.ttf",
            "Arial.ttf",
            "LiberationSans-Regular.ttf",
            "Helvetica.ttf",
            "Verdana.ttf",
        ],
        size=font_size,
    )

    draw = ImageDraw.Draw(barcode_img)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    new_height = height + text_height + padding * 2
    new_image = Image.new("RGB", (width, new_height), "white")
    new_image.paste(barcode_img, (0, 0))

    text_x = max(padding, (width - text_width) // 2)
    text_y = height + padding
    ImageDraw.Draw(new_image).text((text_x, text_y), text, font=font, fill="black")
    return new_image


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


def _exception_details(exc: BaseException) -> Tuple[str, str]:
    """Return a ``(message, stack)`` tuple for the provided exception."""

    message = f"{exc.__class__.__name__}: {exc}"
    stack = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    return message, stack


def _is_font_resource_error(exc: BaseException) -> bool:
    """Return ``True`` when the exception chain indicates a missing font resource."""

    target = "cannot open resource"
    to_check: List[BaseException] = [exc]
    seen: Set[int] = set()

    while to_check:
        current = to_check.pop()
        identifier = id(current)
        if identifier in seen:
            continue
        seen.add(identifier)

        if target in str(current).lower():
            return True

        if getattr(current, "__cause__", None) is not None:
            to_check.append(current.__cause__)  # type: ignore[arg-type]
        if getattr(current, "__context__", None) is not None:
            to_check.append(current.__context__)  # type: ignore[arg-type]

    return False


def _verify_barcode_dependencies() -> Optional[str]:
    """Return ``None`` if barcode dependencies are available, else an error message."""

    dependency_hint = "Install Pillow, python-barcode, and reportlab if needed."

    if _BARCODE_IMPORT_ERROR is not None:
        message, _stack = _exception_details(_BARCODE_IMPORT_ERROR)
        return f"python-barcode unavailable ({message}). {dependency_hint}"

    if barcode is None or ImageWriter is None:  # pragma: no cover - defensive guard
        return f"python-barcode ImageWriter is not available. {dependency_hint}"

    try:
        ImageWriter()
    except Exception as exc:  # pragma: no cover - initialization edge case
        message, _stack = _exception_details(exc)
        return f"Unable to initialize python-barcode ImageWriter ({message}). {dependency_hint}"

    return None


def _select_rinven_barcode_formats(_data: str) -> List[str]:
    """Return the barcode symbologies to attempt for Rinven tags.

    Rinven stakeholders require that all barcodes are rendered using Code 128.
    Historically, the code attempted to optimise by preferring other formats
    such as EAN or UPC for numeric data; however, that behaviour caused the
    generated labels to deviate from the mandated Code 128 symbology. The
    implementation now always returns a single-entry list with ``code128`` so
    that every rendered barcode adheres to this requirement.
    """

    return ["code128"]


def _render_barcode_image(
    data: str,
    bc_format: str,
    writer_options: Optional[Dict[str, object]] = None,
) -> Tuple[Image.Image, _SafeImageWriter]:
    """Render a barcode image using the safe writer helper."""

    if barcode is None or ImageWriter is None:  # pragma: no cover - defensive guard
        raise RuntimeError("Barcode dependencies are unavailable")

    barcode_class = barcode.get_barcode_class(bc_format)

    def _render_with_writer(
        active_writer: _SafeImageWriter,
        options: Optional[Dict[str, object]] = None,
    ) -> Tuple[Image.Image, _SafeImageWriter]:
        instance = barcode_class(data, writer=active_writer)
        img = instance.render(writer_options=options)
        return img, active_writer

    effective_writer_options: Dict[str, object] = {
        "write_text": False,
        "module_width": 1.2,
        "module_height": 10.0,
        "quiet_zone": 2.0,
    }

    if writer_options:
        effective_writer_options.update(writer_options)

    try:
        return _render_with_writer(_SafeImageWriter(), effective_writer_options)
    except Exception as exc:  # noqa: BLE001 - barcode library raised an error
        if not _is_font_resource_error(exc):
            raise

        fallback_options = dict(effective_writer_options)
        fallback_options["write_text"] = False

        fallback_writer = _SafeImageWriter()
        try:
            barcode_img, writer = _render_with_writer(
                fallback_writer, fallback_options
            )
        except Exception:
            raise exc

        barcode_img = _append_barcode_text(barcode_img, data)

        warning = (
            writer.font_warning_message
            or "Barcode font unavailable. Text rendered with fallback font."
        )
        writer.font_fallback_used = True
        writer.font_warning_message = warning
        return barcode_img, writer


def _tighten_barcode_whitespace(
    barcode_img: Image.Image, margin_x: int, margin_y: int
) -> Image.Image:
    """Crop redundant whitespace around a rendered barcode image."""

    if margin_x <= 0 and margin_y <= 0:
        return barcode_img

    grayscale = barcode_img if barcode_img.mode == "L" else barcode_img.convert("L")
    bbox = ImageOps.invert(grayscale).getbbox()
    if not bbox:
        return barcode_img

    left, upper, right, lower = bbox
    left = max(left - margin_x, 0)
    upper = max(upper - margin_y, 0)
    right = min(right + margin_x, barcode_img.width)
    lower = min(lower + margin_y, barcode_img.height)

    if (
        left == 0
        and upper == 0
        and right == barcode_img.width
        and lower == barcode_img.height
    ):
        return barcode_img

    return barcode_img.crop((left, upper, right, lower))


def _render_rinven_barcode(data: str, dpi: int) -> Tuple[Image.Image, str, Optional[str]]:
    """Render the barcode image for Rinven tags.

    Returns a tuple of ``(image, format_name, font_warning_message)`` where the
    warning entry is ``None`` when the preferred TrueType font was available.
    """

    dependency_issue = _verify_barcode_dependencies()
    if dependency_issue is not None:
        raise RuntimeError(dependency_issue)

    errors: List[Dict[str, str]] = []
    for bc_format in _select_rinven_barcode_formats(data):
        try:
            barcode_img, writer = _render_barcode_image(
                data,
                bc_format,
                writer_options={
                    "module_height": 18.0,
                    "quiet_zone": 6,
                },
            )
            warning_message = getattr(writer, "font_warning_message", None)
            return barcode_img, bc_format, warning_message
        except Exception as exc:  # noqa: BLE001 - external library
            message, stack = _exception_details(exc)
            errors.append({"format": bc_format, "message": message, "stack": stack})

    if not errors:
        raise RuntimeError("Barcode rendering failed for an unknown reason")

    # Log all errors for transparency before raising the final exception.
    for entry in errors:
        logging.error(
            "Failed to render %s barcode for Rinven tag: %s\n%s",
            entry.get("format", "unknown"),
            entry.get("message", ""),
            entry.get("stack", ""),
        )

    first_error = errors[0]
    summary = first_error.get("message", "Barcode rendering failed")
    format_name = first_error.get("format")
    if format_name:
        summary = f"{format_name}: {summary}"
    combined_stack = first_error.get("stack", "")
    detailed_message = summary
    if combined_stack:
        detailed_message = f"{summary}\n{combined_stack}"

    raise RuntimeError(detailed_message)


def _normalize_tag_value(value: Optional[str]) -> str:
    """Normalize text for Rinven tag rendering by trimming and collapsing spaces."""

    if value is None:
        return ""

    text = str(value).replace("\u00a0", " ")
    collapsed = " ".join(text.split())
    return collapsed.strip()


def _prepare_rinven_fields(
    details: Dict[str, str], only_filled_fields: bool
) -> Tuple[Dict[str, str], List[Tuple[str, str]], List[str]]:
    """Return normalized details and ordered field/value pairs for rendering."""

    normalized_details: Dict[str, str] = {}
    for key, value in details.items():
        normalized_details[key] = _normalize_tag_value(value)

    field_order: List[Tuple[str, str]] = [
        ("design", "Design"),
        ("color", "Color"),
        ("size", "Size"),
        ("origin", "Origin"),
        ("style", "Style"),
        ("content", "Content"),
        ("type", "Type"),
        ("sku", "SKU"),
        ("rug_no", "Rug #"),
    ]

    included_keys: List[str] = []
    field_entries: List[Tuple[str, str]] = []

    for key, label in field_order:
        value = normalized_details.get(key, "")
        if value:
            field_entries.append((label, value))
            included_keys.append(key)
        elif not only_filled_fields:
            field_entries.append((label, ""))

    return normalized_details, field_entries, included_keys


def build_rinven_tag_image(
    details: Dict[str, str],
    include_barcode: bool,
    barcode_data: Optional[str],
    only_filled_fields: bool = True,
    output_format: str = "png",
    label_size_in: Optional[Tuple[float, float]] = None,
):
    """Render the Rinven tag image and return it with rendering metadata."""

    normalized_details, field_entries, included_keys = _prepare_rinven_fields(
        details, only_filled_fields
    )

    formatted_price = _format_price_text(normalized_details.get("price"))
    if formatted_price:
        normalized_details["price"] = formatted_price

    included_keys = list(included_keys)
    has_price_value = bool(_normalize_tag_value(normalized_details.get("price")))
    if has_price_value and "price" not in included_keys:
        included_keys.insert(0, "price")

    normalized_barcode = _normalize_tag_value(barcode_data)
    barcode_enabled = bool(include_barcode and normalized_barcode)

    warnings: List[Dict[str, str]] = []
    if include_barcode and not normalized_barcode:
        warnings.append({"code": "barcode_missing"})

    is_dymo_output = output_format.lower() == "dymo"
    DPI = 300
    if label_size_in:
        width_in, height_in = label_size_in
        if width_in <= 0 or height_in <= 0:
            raise ValueError("Label dimensions must be positive values.")
    elif is_dymo_output:
        width_in = 2.3
        height_in = 4.0
    else:
        width_in = 4.0
        height_in = 2.3125  # DYMO 30256 Shipping (4" x 2-5/16")
    width_px = int(round(width_in * DPI))
    height_px = int(round(height_in * DPI))
    padding = int(0.12 * DPI)
    top_extra_padding = int(0.06 * DPI)

    try:
        template_image = None
        if not is_dymo_output:
            template_candidates = [
                ("rinven_tag_template.png",),
                ("templates", "rinven_tag_template.png"),
                ("assets", "rinven_tag_template.png"),
            ]
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
                template_image = template_image.resize(
                    (width_px, height_px), Image.Resampling.LANCZOS
                )
            canvas = template_image.copy()
        else:
            if is_dymo_output:
                canvas = Image.new("1", (width_px, height_px), 1)
            else:
                canvas = Image.new("RGB", (width_px, height_px), "white")

        draw = ImageDraw.Draw(canvas)
        text_fill = 0 if is_dymo_output else "black"

        title_font_size = int(0.18 * DPI)
        title_pref_fonts = [
            "arialbd.ttf",
            "Arial Bold.ttf",
            "Helvetica-Bold.ttf",
            "Verdana Bold.ttf",
            "LiberationSans-Bold.ttf",
        ]
        font_size_value = details.get("font_size", "20")
        try:
            font_size_pt = int(str(font_size_value).strip())
        except (TypeError, ValueError):
            font_size_pt = 20
        base_font_px = max(int(round((font_size_pt / 72.0) * DPI)), 1)
        text_font_size = base_font_px
        text_pref_fonts = [
            "arial.ttf",
            "Helvetica.ttf",
            "Verdana.ttf",
            "LiberationSans-Regular.ttf",
        ]

        current_y = padding + top_extra_padding

        price_text = normalized_details.get("price", "")
        price_drawn = False
        if price_text:
            price_pref_fonts = title_pref_fonts
            price_font_size = int(0.18 * DPI)
            min_price_font = max(int(0.1 * DPI), 30)
            price_font, price_font_size = _fit_font(
                draw,
                price_text,
                price_pref_fonts,
                width_px - (padding * 2),
                price_font_size,
                min_price_font,
            )
            bbox = draw.textbbox((0, 0), price_text, font=price_font)
            price_w = bbox[2] - bbox[0]
            price_h = bbox[3] - bbox[1]
            price_x = (width_px - price_w) // 2
            draw.text((price_x, current_y), price_text, fill=text_fill, font=price_font)
            current_y += price_h + int(0.05 * DPI)
            price_drawn = True

        barcode_img: Optional[Image.Image] = None
        barcode_format_used: Optional[str] = None
        if barcode_enabled:
            try:
                raw_barcode, barcode_format_used, font_warning_message = _render_rinven_barcode(
                    normalized_barcode,
                    DPI,
                )
                target_mode = "1" if is_dymo_output else "RGB"
                barcode_img = raw_barcode.convert(target_mode)
                # Preserve a generous quiet zone around the barcode to keep it
                # easily scannable even after the surrounding whitespace is
                # trimmed. Empirically a wider horizontal margin greatly
                # improves reliability for handheld scanners, so reserve more
                # space than the bare minimum before cropping.
                # Expand the quiet zone to keep scanners happy even when the label is
                # trimmed closely or printed slightly off-center.
                margin_x = max(18, int(round(0.1 * DPI)))
                margin_y = max(12, int(round(0.06 * DPI)))
                barcode_img = _tighten_barcode_whitespace(barcode_img, margin_x, margin_y)
                if font_warning_message:
                    has_font_warning = any(
                        isinstance(entry, dict) and entry.get("code") == "barcode_font_warning"
                        for entry in warnings
                    )
                    if not has_font_warning:
                        warnings.append(
                            {
                                "code": "barcode_font_warning",
                                "message": font_warning_message,
                            }
                        )
            except Exception as exc:  # pragma: no cover - GUI feedback
                message, stack = _exception_details(exc)
                logging.error(
                    "Unable to render barcode for Rinven tag: %s\n%s",
                    message,
                    stack,
                )
                warnings.append(
                    {
                        "code": "barcode_error",
                        "message": message,
                        "stack": stack,
                    }
                )

        if barcode_img is not None:
            # Draw the barcode before any text so it remains visually on top of the canvas.
            max_barcode_width = width_px - (padding * 2)
            max_barcode_height = int(height_px * 0.34)
            min_barcode_width = min(
                max_barcode_width,
                max(int(max_barcode_width * 0.9), 360),
            )
            min_barcode_height = min(
                max_barcode_height,
                max(int(max_barcode_height * 0.9), 180),
            )

            width = max(barcode_img.width, 1)
            height = max(barcode_img.height, 1)

            scale_up = max(
                min_barcode_width / width,
                min_barcode_height / height,
                1.0,
            )
            if scale_up > 1.0:
                new_size = (
                    int(round(width * scale_up)),
                    int(round(height * scale_up)),
                )
                # Upscaling with anti-aliasing blurs the barcode bars; use NEAREST to keep
                # module edges crisp and scannable.
                barcode_img = barcode_img.resize(new_size, _RESAMPLE_NEAREST)
                width, height = barcode_img.size

            scale_down = min(
                max_barcode_width / max(width, 1),
                max_barcode_height / max(height, 1),
                1.0,
            )
            if scale_down < 1.0:
                new_size = (
                    int(round(width * scale_down)),
                    int(round(height * scale_down)),
                )
                # Downscaling with NEAREST avoids introducing grey artifacts between bars.
                barcode_img = barcode_img.resize(new_size, _RESAMPLE_NEAREST)

            barcode_x = (width_px - barcode_img.width) // 2
            canvas.paste(barcode_img, (barcode_x, current_y))
            current_y += barcode_img.height + int(0.05 * DPI)

        collection_name = normalized_details.get("collection", "")
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
            draw.text((title_x, current_y), collection_name, fill=text_fill, font=title_font)
            current_y += title_h + int(0.07 * DPI)

        field_padding_top = max(current_y, padding)
        current_y = field_padding_top

        base_colon_gap = int(0.03 * DPI)
        base_value_gap = int(0.05 * DPI)
        available_line_width = width_px - (padding * 2)
        available_height = height_px - padding - current_y
        min_text_size = max(int(text_font_size * 0.45), 14)

        max_label_width = colon_width = line_height = colon_gap = value_gap = 0
        max_line_width = 0

        if field_entries:

            def gap_values(font_size):
                local_colon_gap = max(base_colon_gap, int(font_size * 0.2))
                local_value_gap = max(base_value_gap, int(font_size * 0.4))
                return local_colon_gap, local_value_gap

            def measurements(font, font_size):
                local_colon_gap, local_value_gap = gap_values(font_size)
                colon_bbox = draw.textbbox((0, 0), ":", font=font)
                local_colon_width = colon_bbox[2] - colon_bbox[0]
                local_line_height = int(font_size * 1.35)
                widest_label = 0
                local_max_line_width = 0
                fits_width = True
                for label, value in field_entries:
                    label_bbox = draw.textbbox((0, 0), label, font=font)
                    label_width = label_bbox[2] - label_bbox[0]
                    value_text = value.strip()
                    value_bbox = draw.textbbox((0, 0), value_text, font=font)
                    value_width = value_bbox[2] - value_bbox[0]
                    line_width = (
                        label_width
                        + local_colon_gap
                        + local_colon_width
                        + local_value_gap
                        + value_width
                    )
                    local_max_line_width = max(local_max_line_width, line_width)
                    if line_width > available_line_width:
                        fits_width = False
                        break
                    widest_label = max(widest_label, label_width)
                return (
                    fits_width,
                    widest_label,
                    local_colon_width,
                    local_line_height,
                    local_colon_gap,
                    local_value_gap,
                    local_max_line_width,
                )

            text_font = _load_font(text_pref_fonts, size=text_font_size)
            (
                fits_width,
                max_label_width,
                colon_width,
                line_height,
                colon_gap,
                value_gap,
                max_line_width,
            ) = measurements(text_font, text_font_size)
            total_height = line_height * len(field_entries)

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
                    max_line_width,
                ) = measurements(text_font, text_font_size)
                total_height = line_height * len(field_entries)

            if not fits_width or total_height > available_height:
                longest_value = max(
                    (value.strip() for _, value in field_entries),
                    key=len,
                    default="",
                )
                text_font, text_font_size = _fit_font(
                    draw,
                    longest_value,
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
                    max_line_width,
                ) = measurements(text_font, text_font_size)
                total_height = line_height * len(field_entries)

            text_block_width = max_line_width
            start_x = max(padding, (width_px - text_block_width) // 2)

            for label, value in field_entries:
                label_text = label
                value_text = value.strip()
                draw.text((start_x, current_y), label_text, fill=text_fill, font=text_font)
                colon_x = start_x + max_label_width + colon_gap
                draw.text((colon_x, current_y), ":", fill=text_fill, font=text_font)
                value_x = colon_x + colon_width + value_gap
                draw.text((value_x, current_y), value_text, fill=text_fill, font=text_font)
                current_y += line_height

        has_content = bool(
            price_drawn
            or collection_name
            or (barcode_img is not None)
            or any(value for _, value in field_entries)
        )

        metadata = {
            "normalized_details": normalized_details,
            "included_field_keys": included_keys,
            "warnings": warnings,
            "barcode_used": bool(barcode_img),
            "barcode_format": barcode_format_used,
            "has_content": has_content,
            "has_price": price_drawn,
        }

        return canvas, metadata

    except Exception:
        logging.exception("Error rendering Rinven tag")
        raise


def generate_rinven_tag_label(
    details,
    fname,
    include_barcode,
    barcode_data,
    only_filled_fields: bool = True,
    output_format: str = "png",
    label_size_in: Optional[Tuple[float, float]] = None,
):
    """Creates a Rinven tag label sized for a portrait 4" x 6" label."""

    try:
        is_dymo_output = output_format.lower() == "dymo"
        canvas, metadata = build_rinven_tag_image(
            details,
            include_barcode,
            barcode_data,
            only_filled_fields=only_filled_fields,
            output_format=output_format,
            label_size_in=label_size_in,
        )

        if not metadata.get("has_content"):
            return ("⚠️ No content available to export.", None, metadata)

        if is_dymo_output:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".bmp") as tmp_file:
                saved_path = Path(tmp_file.name)
            canvas.save(saved_path, format="BMP")
        else:
            canvas.save(fname)
            saved_path = Path(fname)
        metadata = dict(metadata or {})
        metadata["output_path"] = str(saved_path)
        return (
            f"✅ Rinven tag saved as '{saved_path}'",
            f"Rinven tag saved to:\n{saved_path.resolve()}",
            metadata,
        )
    except Exception as exc:  # pragma: no cover - GUI feedback
        logging.exception("Error generating Rinven tag")
        return (f"Error generating Rinven tag: {exc}", None, {"warnings": ["render_error"]})


def send_image_to_printer(printer_name: str, image_path: str) -> None:
    """Send the specified image to a system printer."""

    if not printer_name:
        raise RuntimeError("Printer name is required.")
    if not image_path or not os.path.exists(image_path):
        raise RuntimeError(f"File not found: {image_path}")

    system = platform.system()
    if system == "Windows":
        if win32print is None:
            raise RuntimeError("win32print module is not available.")

        import win32ui  # type: ignore
        from PIL import ImageWin

        handle = win32print.OpenPrinter(printer_name)
        hdc = None
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)

            printable_area = (
                hdc.GetDeviceCaps(8),  # HORZRES
                hdc.GetDeviceCaps(10),  # VERTRES
            )
            printer_size = (
                hdc.GetDeviceCaps(110),  # PHYSICALWIDTH
                hdc.GetDeviceCaps(111),  # PHYSICALHEIGHT
            )

            image = Image.open(image_path)
            if image.mode not in {"RGB", "RGBA"}:
                image = image.convert("RGB")

            scale = min(
                printable_area[0] / image.width if image.width else 1,
                printable_area[1] / image.height if image.height else 1,
            )
            scale = min(scale, 1.0)

            target_size = (
                max(1, int(image.width * scale)),
                max(1, int(image.height * scale)),
            )
            if image.size != target_size:
                image = image.resize(target_size, resample=Image.LANCZOS)

            left = int((printer_size[0] - target_size[0]) / 2)
            top = int((printer_size[1] - target_size[1]) / 2)
            right = left + target_size[0]
            bottom = top + target_size[1]

            hdc.StartDoc("Rinven Tag")
            hdc.StartPage()
            dib = ImageWin.Dib(image)
            dib.draw(hdc.GetHandleOutput(), (left, top, right, bottom))
            hdc.EndPage()
            hdc.EndDoc()
        finally:
            if hdc is not None:
                try:
                    hdc.DeleteDC()
                except Exception:
                    pass
            win32print.ClosePrinter(handle)
    else:
        if cups is None:
            raise RuntimeError("cups module is not available.")
        connection = cups.Connection()
        connection.printFile(printer_name, image_path, "Rinven Tag", {})


def generate_bulk_rinven_tags(
    file_path: str,
    output_dir: str,
    include_barcode: bool,
    only_filled_fields: bool,
    font_size_value: Optional[str],
    output_format: str,
    label_size_in: Optional[Tuple[float, float]],
    log_callback=None,
) -> Tuple[List[Tuple[Dict[str, str], str]], str, str]:
    """Generate Rinven tags in bulk and return the created files with their details.

    The returned tuple contains ``(generated_files, summary_message, status)`` where
    ``generated_files`` is a list of ``(details, output_path)`` tuples suitable for
    downstream printing, ``summary_message`` is a human readable description of the
    run, and ``status`` is one of ``"Success"`` or ``"Warning"``.
    """

    def _log(message: str) -> None:
        if log_callback is None:
            return
        try:
            log_callback(message)
        except Exception:  # pragma: no cover - defensive logging
            pass

    source = Path(file_path)
    if not source.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    render_format = (output_format or "png").lower()
    normalized_label_size = label_size_in if isinstance(label_size_in, tuple) else None

    try:
        output_path = Path(output_dir) if output_dir else source.parent / "rinven_tags"
        output_path.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise RuntimeError(f"Could not create output folder: {exc}") from exc

    _log(f"Starting Rinven tag generation from: {source}")

    try:
        if source.suffix.lower() in {".xlsx", ".xls"}:
            dataframe = pd.read_excel(source, dtype=str)
        else:
            dataframe = pd.read_csv(source, dtype=str)
    except Exception as exc:
        raise RuntimeError(f"Excel/CSV file could not be read: {exc}") from exc

    if dataframe.empty:
        return [], "The selected file does not contain any rows.", "Warning"

    column_lookup = {str(column).strip().lower(): column for column in dataframe.columns}

    def _coerce_font_size(raw_value: Optional[str]) -> Optional[str]:
        try:
            cleaned = str(raw_value).strip()
        except Exception:
            return None

        if not cleaned:
            return None

        try:
            value = int(cleaned)
        except (TypeError, ValueError):
            return None
        return str(max(value, 1))

    default_font_size = _coerce_font_size(font_size_value)

    total_rows = len(dataframe)
    generated = 0
    skipped = 0
    failures = 0
    generated_files: List[Tuple[Dict[str, str], str]] = []
    slug_counts: Dict[str, int] = {}

    for index, row in dataframe.iterrows():
        details = {
            "collection": _pick_first_value(row, column_lookup, ["collection", "vcollection", "collection name"]),
            "design": _pick_first_value(row, column_lookup, ["design", "rugno", "rug no", "rug #"]),
            "size": _pick_first_value(row, column_lookup, ["size", "asize", "stsize"]),
            "origin": _pick_first_value(row, column_lookup, ["origin"]),
            "style": _pick_first_value(row, column_lookup, ["style"]),
            "content": _pick_first_value(row, column_lookup, ["content", "material", "materials"]),
            "type": _pick_first_value(row, column_lookup, ["type"]),
            "rug_no": _pick_first_value(row, column_lookup, ["rug #", "rug no", "rugno", "design"]),
            "sku": _pick_first_value(row, column_lookup, ["sku", "upc", "barcode", "item sku"]),
        }

        color_value = _pick_first_value(row, column_lookup, ["color"])
        if not color_value:
            ground = _pick_first_value(row, column_lookup, ["ground"])
            border = _pick_first_value(row, column_lookup, ["border"])
            if ground and border:
                color_value = f"{ground}/{border}"
            else:
                color_value = ground or border
        details["color"] = color_value

        price_raw = _pick_first_value(row, column_lookup, ["price", "retail", "amount", "sp", "msrp"])
        details["price"] = _format_price_text(price_raw)

        row_font_size = _coerce_font_size(
            _pick_first_value(row, column_lookup, ["font_size", "font size", "font size (pt)"])
        )
        font_size = row_font_size or default_font_size
        if font_size:
            details["font_size"] = font_size

        barcode_value = _pick_first_value(row, column_lookup, ["barcode", "upc", "sku", "rugno"])
        use_barcode = bool(include_barcode and barcode_value)

        try:
            canvas, metadata = build_rinven_tag_image(
                details,
                use_barcode,
                barcode_value,
                only_filled_fields=only_filled_fields,
                output_format=render_format,
                label_size_in=normalized_label_size,
            )
        except Exception as exc:
            failures += 1
            _log(f"Error rendering row {index + 1}: {exc}")
            continue

        if not metadata.get("has_content"):
            skipped += 1
            _log(f"Skipping row {index + 1}: No content available for tag.")
            continue

        slug_source = (
            details.get("rug_no")
            or details.get("design")
            or details.get("sku")
            or f"row-{index + 1}"
        )
        slug = _slugify_tag_filename(slug_source, f"row-{index + 1}")
        slug_occurrence = slug_counts.get(slug, 0)
        slug_counts[slug] = slug_occurrence + 1
        if slug_occurrence:
            new_slug = f"{slug}-{slug_occurrence + 1}"
            _log(
                f"Duplicate filename slug '{slug}' on row {index + 1}; saving as '{new_slug}'."
            )
            slug = new_slug
        output_file = output_path / f"rinven_tag_{slug}.png"

        try:
            canvas.save(output_file)
        except Exception as exc:
            failures += 1
            _log(f"Failed to save tag for row {index + 1}: {exc}")
            continue

        job_details = dict(details)
        job_details["_include_barcode"] = use_barcode
        job_details["_barcode_value"] = barcode_value
        job_details["_only_filled_fields"] = only_filled_fields
        if normalized_label_size:
            job_details["_label_size_in"] = normalized_label_size

        generated_files.append((job_details, str(output_file)))
        generated += 1

        progress_interval = max(1, total_rows // 10)
        if (index + 1) % progress_interval == 0:
            percentage = ((index + 1) / total_rows) * 100
            _log(f"  ...Progress: {percentage:.0f}% ({index + 1}/{total_rows})")

    summary_parts = [f"Generated {generated} tag(s)."]
    if skipped:
        summary_parts.append(f"Skipped {skipped} row(s) with no content.")
    if failures:
        summary_parts.append(f"Encountered {failures} error(s).")
    summary_parts.append(f"Files saved to: {output_path}")

    status = "Success" if failures == 0 else "Warning"
    summary_message = "\n".join(summary_parts)
    _log(summary_message)

    return generated_files, summary_message, status


def print_bulk_rinven_tags(
    printer_name: str,
    generated_files: List[Tuple[Dict[str, str], str]],
    font_size_pt: Optional[str],
    progress_callback,
) -> None:
    """Render and print the provided Rinven tag details to a thermal printer."""

    if not generated_files:
        return

    try:
        override_font_size = int(str(font_size_pt).strip()) if font_size_pt else None
    except (TypeError, ValueError):  # pragma: no cover - defensive parsing
        override_font_size = None

    total = len(generated_files)
    for index, (details, path) in enumerate(generated_files, start=1):
        context_details = dict(details or {})
        include_barcode = bool(
            context_details.pop("_include_barcode", context_details.get("include_barcode", False))
        )
        barcode_value = context_details.pop("_barcode_value", None)
        if barcode_value is None:
            barcode_value = context_details.get("barcode_value") or context_details.get("barcode")
        only_filled_fields = bool(context_details.pop("_only_filled_fields", True))
        label_size_in = context_details.pop("_label_size_in", None)

        if override_font_size:
            context_details.setdefault("font_size", str(override_font_size))

        canvas, _ = build_rinven_tag_image(
            context_details,
            include_barcode,
            barcode_value,
            only_filled_fields=only_filled_fields,
            output_format="dymo",
            label_size_in=label_size_in,
        )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".bmp") as tmp:
            temp_path = tmp.name
            canvas.save(temp_path, format="BMP")

        try:
            send_image_to_printer(printer_name, temp_path)
        finally:
            try:
                os.remove(temp_path)
            except OSError:
                pass

        if progress_callback:
            try:
                progress_callback(index, total, os.path.basename(path))
            except Exception:  # pragma: no cover - UI callback safety
                pass


def _pick_first_value(row, column_lookup, candidates: Iterable[str]) -> str:
    """Return the first non-empty candidate value from a row using case-insensitive names."""

    for candidate in candidates:
        key = candidate.strip().lower()
        column_name = column_lookup.get(key)
        if column_name is None:
            continue
        try:
            value = row[column_name]
        except KeyError:
            continue
        if pd.isna(value):
            continue
        text = _normalize_tag_value(str(value))
        if text:
            return text
    return ""


def _format_price_text(raw: str) -> str:
    """Normalise a raw price string into a ``Price $x.xx`` format when possible."""

    text = _normalize_tag_value(raw)
    if not text:
        return ""

    prefix_pattern = re.compile(r"^price[:\s\$]*", re.IGNORECASE)
    stripped = prefix_pattern.sub("", text)

    cleaned = stripped.replace("$", "").replace(",", "")
    try:
        value = float(cleaned)
    except ValueError:
        return text if prefix_pattern.match(text) else f"Price {text}" if text else ""

    formatted_value = f"{value:,.2f}".rstrip("0").rstrip(".")
    return f"Price ${formatted_value}"


def _slugify_tag_filename(text: str, fallback: str) -> str:
    """Return a filesystem-safe slug for generated Rinven tag filenames."""

    normalized = _normalize_tag_value(text)
    if normalized:
        slug = re.sub(r"[^A-Za-z0-9]+", "-", normalized).strip("-")
        if slug:
            return slug[:80]
    return fallback


def generate_rinven_tags_from_file_task(
    file_path: str,
    output_dir: str,
    include_barcode: bool,
    only_filled_fields: bool,
    font_size_value: Optional[str],
    output_format: str,
    label_size_in: Optional[Tuple[float, float]],
    log_callback,
    completion_callback,
):
    """Generate Rinven tag images for each row of an Excel/CSV file."""
    try:
        generated_files, summary_message, status = generate_bulk_rinven_tags(
            file_path,
            output_dir,
            include_barcode,
            only_filled_fields,
            font_size_value,
            output_format,
            label_size_in,
            log_callback,
        )
    except Exception as exc:
        log_callback(f"Error: {exc}")
        completion_callback("Error", str(exc))
        return

    completion_callback(status, summary_message)

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
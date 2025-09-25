# backend_logic.py
import os
import re
import logging
import shutil
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pillow_heif
import qrcode
import barcode
from barcode.writer import ImageWriter
# tqdm kütüphanesi artık kullanılmadığı için import satırını silebiliriz veya bırakabiliriz.
# from tqdm import tqdm

os.environ['GRPC_DNS_RESOLVER'] = 'native'
import google.generativeai as genai

# This file contains NO tkinter code.

# --- Helper Functions ---

def clean_file_path(file_path: str) -> str:
    """Removes quotes and whitespace from a file path string."""
    return file_path.strip().strip('"').strip("'")

def parse_feet_inches(value_str: str):
    """Parses various string formats (e.g., 5'2", 5.2', 8") into a decimal foot value."""
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
    """Converts a dimension string (e.g., "5'2" x 8'") into a tuple of (width_in, height_in)."""
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
    """Creates a label image for Dymo printers with the code centered and optional text."""
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

def initialize_gemini_model(api_key):
    """Configures the API and returns the model object, or an error."""
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        return model, None
    except Exception as e:
        return None, e

def ask_ai(model, prompt):
    """Sends a prompt to the configured Gemini model."""
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Sorry, an error occurred: {e}"

def process_files_task(src, tgt, nums_f, action, log_callback, completion_callback):
    """Finds files based on a list and copies or moves them."""
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
    
    # tqdm'i GUI log alanına yönlendirmek kilitlenmeye neden olabildiği için
    # basit bir manuel ilerleme bildirimi kullanıyoruz.
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
    """Resizes all images in a folder based on width or percentage."""
    tgt_folder = os.path.join(src_folder, "resized")
    os.makedirs(tgt_folder, exist_ok=True)
    log_callback(f"Resized images will be saved in: {tgt_folder}")
    files = [f for f in os.listdir(src_folder) if f.lower().endswith(('.jpg','.jpeg','.png','.webp'))]
    if not files:
        log_callback("No compatible images found."); return
    
    log_callback(f"Starting resize process for {len(files)} images...")
    for i, f in enumerate(files):
        if (i + 1) % 20 == 0: # Her 20 resimde bir ilerleme bildir
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
    """Safely processes a single rug dimension string."""
    try:
        w_in, h_in = size_to_inches_wh(s)
        area = calculate_sqft(s)
        return {'w': w_in, 'h': h_in, 'a': area}
    except Exception:
        return {'w': None, 'h': None, 'a': None}

def bulk_rug_sizer_task(path, col, log_callback, completion_callback):
    """Processes a sheet of rug dimensions and adds calculated columns."""
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

    # tqdm ve progress_apply yerine manuel, daha güvenli bir döngü kullanıyoruz
    total_rows = len(df)
    results = []
    for i, row_value in enumerate(df[sel_col]):
        results.append(_process_rug_size_row(row_value))
        # Her %10'da bir veya her 100 satırda bir (hangisi daha sık ise) ilerleme bildir
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
        log_callback(f"Could not save as Excel ({e}). ✅ Saved to CSV instead: {csv_path}")
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

def add_image_links_task(input_path, links_path, key_col, log_callback, completion_callback):
    """
    Anahtar bir sütuna göre, resim bağlantılarını ayrı bir CSV dosyasından
    bir Excel/CSV dosyasına ekler.
    """
    log_callback("Resim bağlantılarını ekleme işlemi başlatılıyor...")
    try:
        # Ana veri dosyasını yükle (Excel veya CSV)
        if input_path.lower().endswith((".xlsx", ".xls")):
            df_main = pd.read_excel(input_path)
        else:
            df_main = pd.read_csv(input_path)
            
        # Anahtar sütunu bul
        sel_col = None
        if len(key_col) == 1 and key_col.isalpha():
            idx = ord(key_col.upper()) - ord('A')
            if idx < len(df_main.columns):
                sel_col = df_main.columns[idx]
        elif key_col in df_main.columns:
            sel_col = key_col
            
        if not sel_col:
            completion_callback("Hata", f"Anahtar sütun '{key_col}' bulunamadı.")
            return

        # Resim bağlantıları dosyasını yükle
        log_callback("Resim bağlantıları dosyası yükleniyor...")
        df_links = pd.read_csv(links_path, header=None, dtype=str)
        
        # Bağlantıları hızlı arama için bir sözlükte grupla
        link_map = {}
        for link in df_links[0].dropna().tolist():
            # Anahtar numarayı (örneğin, "073910") çıkar ve bağlantıları grupla
            match = re.search(r"/(\d{6,})(-\d+)?\.jpg", link)
            if match:
                key = match.group(1)
                if key not in link_map:
                    link_map[key] = []
                link_map[key].append(link)
        
        # Tutarlılık için bağlantıları sırala (-1, -2, vb.)
        for key in link_map:
            link_map[key].sort()

        # Ana DataFrame'de dolaşarak bağlantıları ekle
        log_callback("Bağlantılar eşleştiriliyor ve veriye ekleniyor...")
        
        for index, row in df_main.iterrows():
            key_val = str(row[sel_col]).strip()
            if key_val in link_map:
                links = link_map[key_val]
                # Bağlantılar için yeni sütunlar ekle
                for i, link in enumerate(links):
                    col_name = f"Image_Link_{i + 1}"
                    df_main.loc[index, col_name] = link
            else:
                log_callback(f"Anahtar '{key_val}' için resim bağlantısı bulunamadı.")
                
        # Güncellenmiş dosyayı kaydet
        out_path = f"{os.path.splitext(input_path)[0]}_with_images.xlsx"
        try:
            df_main.to_excel(out_path, index=False)
            log_callback(f"✅ Dosya başarıyla kaydedildi: {out_path}")
            completion_callback("Başarılı", f"Bağlantılar eklendi ve dosya şu konuma kaydedildi:\n{out_path}")
        except Exception as e:
            csv_path = f"{os.path.splitext(input_path)[0]}_with_images.csv"
            df_main.to_csv(csv_path, index=False)
            log_callback(f"Excel olarak kaydedilemedi ({e}). ✅ Bunun yerine CSV olarak kaydedildi: {csv_path}")
            completion_callback("CSV Olarak Kaydedildi", f"Excel olarak kaydedilemedi. Bunun yerine CSV olarak kaydedildi:\n{csv_path}")

    except Exception as e:
        log_callback(f"Bir hata oluştu: {e}")
        completion_callback("Hata", f"Bir hata oluştu: {e}")

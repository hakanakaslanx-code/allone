# backend_logic.py
import os
import re
import json
import logging
import shutil
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pillow_heif
import qrcode
import barcode
from barcode.writer import ImageWriter
import google.generativeai as genai

# This file contains NO tkinter code.

def clean_file_path(file_path: str) -> str:
    return file_path.strip().strip('"').strip("'")

# ... (parse_feet_inches, size_to_inches_wh, calculate_sqft gibi tüm yardımcı fonksiyonlar buraya gelecek)
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


# --- Task Functions ---

def ask_ai(model, prompt):
    """Sends a prompt to the configured Gemini model."""
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Sorry, an error occurred: {e}"

def resize_images_task(src_folder, mode, value, quality, log_callback, completion_callback):
    from tqdm import tqdm
    tgt_folder = os.path.join(src_folder, "resized")
    os.makedirs(tgt_folder, exist_ok=True)
    log_callback(f"Resized images will be saved in: {tgt_folder}")
    files = [f for f in os.listdir(src_folder) if f.lower().endswith(('.jpg','.jpeg','.png','.webp'))]
    if not files:
        log_callback("No compatible images found.")
        return
    log_callback(f"Starting resize process in '{mode}' mode...")
    for f in tqdm(files, desc="Resizing images"):
        try:
            with Image.open(os.path.join(src_folder, f)) as img:
                ow, oh = img.size; nw, nh = ow, oh
                if mode == "width":
                    if ow > value:
                        ratio = value / float(ow); nw, nh = value, int(float(oh) * ratio)
                else: # percentage
                    nw, nh = int(ow * value / 100), int(oh * value / 100)
                if nw < 1: nw = 1
                if nh < 1: nh = 1
                if (nw, nh) != (ow, oh):
                    resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    img = img.resize((nw, nh), resample)
                if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                save_path = os.path.join(tgt_folder, f)
                if f.lower().endswith(('.jpg','.jpeg')):
                    img.save(save_path, "JPEG", quality=quality, optimize=True)
                else:
                    img.save(save_path)
                log_callback(f"Resized: {f} -> {nw}x{nh}")
        except Exception as e:
            log_callback(f"Error with {f}: {e}")
    log_callback("\n✅ Image resizing complete.")
    completion_callback("Success", "Image processing is complete.")

# ... Diğer tüm görev fonksiyonları (process_files, convert_heic vb.)
# benzer şekilde, log_callback ve completion_callback alacak şekilde buraya eklenebilir.

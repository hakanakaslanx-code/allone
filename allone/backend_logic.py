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
from tqdm import tqdm
import google.generativeai as genai

# This file contains NO tkinter code.

# --- Helper Functions ---

def clean_file_path(file_path: str) -> str:
    """Removes quotes and whitespace from a file path string."""
    return file_path.strip().strip('"').strip("'")
# ... (Diğer tüm yardımcı fonksiyonlar aynı kalacak, sadece en alta yeni bir fonksiyon ekledim)
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

def convert_units_logic(input_string):
    i = input_string.lower()
    if not i: return ""
    m = re.match(r"^\s*(.+?)\s*(cm|m|ft|in)\s+to\s+(cm|m|ft|in)\s*$", i, re.I)
    if not m: return "Invalid Format"
    v_str, fu, tu = m.groups(); cm = None
    try:
        if fu == 'ft': cm = parse_feet_inches(v_str) * 30.48 if parse_feet_inches(v_str) else None
        else: val = float(v_str); cm = val if fu == 'cm' else val * 100 if fu == 'm' else val * 2.54 if fu == 'in' else None
    except: pass
    if cm is None: return f"Could not parse '{v_str}'."
    res = ""
    if tu == 'cm': res = f"{cm:.2f} cm"
    elif tu == 'm': res = f"{cm / 100:.2f} m"
    elif tu == 'in': res = f"{cm / 2.54:.2f} in"
    elif tu == 'ft': total_in = cm / 2.54; res = f"{int(total_in // 12)}' {total_in % 12:.2f}\""
    return f"--> {res}"

# --- Task Functions ---

def initialize_gemini_model(api_key):
    """Configures the API and returns the model object, or an error."""
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        return model, None  # (model, error)
    except Exception as e:
        return None, e

def ask_ai(model, prompt):
    """Sends a prompt to the configured Gemini model."""
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Sorry, an error occurred: {e}"

# (Diğer tüm backend fonksiyonları burada devam ediyor, onlar aynı kalabilir)
def process_files_task(src, tgt, nums_f, action, log_callback, completion_callback):
    # ...
    pass
# ...

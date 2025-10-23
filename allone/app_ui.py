"""Main tkinter user interface for the desktop utility application."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os

from settings_manager import load_settings, save_settings
from updater import check_for_updates
import backend_logic as backend

__version__ = "3.5.3"

TRANSLATIONS = {
    "en": {
        "Combined Utility Tool": "Combined Utility Tool",
        "Welcome to the Combined Utility Tool!": "Welcome to the Combined Utility Tool!",
        "File & Image Tools": "File & Image Tools",
        "Data & Calculation": "Data & Calculation",
        "Code Generators": "Code Generators",
        "Help & About": "Help & About",
        "Language": "Language",
        "English": "English",
        "Turkish": "Turkish",
        "Language changed to {language}.": "Language changed to {language}.",
        "1. Copy/Move Files by List": "1. Copy/Move Files by List",
        "Source Folder:": "Source Folder:",
        "Target Folder:": "Target Folder:",
        "Numbers File (List):": "Numbers File (List):",
        "Browse...": "Browse...",
        "Copy Files": "Copy Files",
        "Move Files": "Move Files",
        "Save Settings": "Save Settings",
        "2. Convert HEIC to JPG": "2. Convert HEIC to JPG",
        "Folder with HEIC files:": "Folder with HEIC files:",
        "Convert": "Convert",
        "3. Batch Image Resizer": "3. Batch Image Resizer",
        "Image Folder:": "Image Folder:",
        "Resize Mode:": "Resize Mode:",
        "By Width": "By Width",
        "By Percentage": "By Percentage",
        "Max Width:": "Max Width:",
        "Percentage (%):": "Percentage (%):",
        "JPEG Quality (1-95):": "JPEG Quality (1-95):",
        "Resize & Compress": "Resize & Compress",
        "4. Format Numbers from File": "4. Format Numbers from File",
        "Excel/CSV/TXT File:": "Excel/CSV/TXT File:",
        "Format": "Format",
        "5. Rug Size Calculator (Single)": "5. Rug Size Calculator (Single)",
        "Dimension (e.g., 5'2\" x 8'):": "Dimension (e.g., 5'2\" x 8'):",
        "Calculate": "Calculate",
        "6. BULK Process Rug Sizes from File": "6. BULK Process Rug Sizes from File",
        "Excel/CSV File:": "Excel/CSV File:",
        "Column Name/Letter:": "Column Name/Letter:",
        "Process File": "Process File",
        "7. Unit Converter": "7. Unit Converter",
        "Conversion:": "Conversion:",
        "182 cm to ft": "182 cm to ft",
        "8. Match Image Links": "8. Match Image Links",
        "Source Excel/CSV File:": "Source Excel/CSV File:",
        "Image Links File (CSV):": "Image Links File (CSV):",
        "Key Column Name/Letter:": "Key Column Name/Letter:",
        "Match and Add Links": "Match and Add Links",
        "8. QR Code Generator": "8. QR Code Generator",
        "Data/URL:": "Data/URL:",
        "Output Type:": "Output Type:",
        "Standard PNG": "Standard PNG",
        "Dymo Label": "Dymo Label",
        "Dymo Size:": "Dymo Size:",
        "Bottom Text:": "Bottom Text:",
        "Filename:": "Filename:",
        "Generate QR Code": "Generate QR Code",
        "9. Barcode Generator": "9. Barcode Generator",
        "Data:": "Data:",
        "Format:": "Format:",
        "Output Type:": "Output Type:",
        "Generate Barcode": "Generate Barcode",
        "Check for Updates": "Check for Updates",
        "Warning": "Warning",
        "Source and Target folders cannot be empty.": "Source and Target folders cannot be empty.",
        "✅ Settings saved to settings.json": "✅ Settings saved to settings.json",
        "Success": "Success",
        "Folder settings have been saved.": "Folder settings have been saved.",
        "Error": "Error",
        "Please specify Source, Target, and Numbers File.": "Please specify Source, Target, and Numbers File.",
        "Please select a valid folder.": "Please select a valid folder.",
        "Please select a valid image folder.": "Please select a valid image folder.",
        "Resize values and quality must be valid numbers.": "Resize values and quality must be valid numbers.",
        "Please select a file.": "Please select a file.",
        "Please enter a dimension.": "Please enter a dimension.",
        "Invalid Format": "Invalid Format",
        "W: {width} in | H: {height} in | Area: {area} sqft": "W: {width} in | H: {height} in | Area: {area} sqft",
        "Please select a file and specify a column.": "Please select a file and specify a column.",
        "Please fill in all file paths and the column name.": "Please fill in all file paths and the column name.",
        "Data and filename are required.": "Data and filename are required.",
        "Error: {message}": "Error: {message}",
        "ABOUT_CONTENT": (
            "Combined Utility Tool - v{version}\n"
            "This application combines common file, image, and data processing tasks into a single interface.\n"
            "--- FEATURES ---\n"
            "1. Copy/Move Files by List:\n"
            "   Finds and copies or moves image files based on a list in an Excel or text file.\n"
            "2. Convert HEIC to JPG:\n"
            "   Converts Apple's HEIC format images to the universal JPG format.\n"
            "3. Batch Image Resizer:\n"
            "   Resizes images by a fixed width or by a percentage of the original dimensions.\n"
            "4. Format Numbers from File:\n"
            "   Formats items from a file's first column into a single, comma-separated line.\n"
            "5. Rug Size Calculator (Single):\n"
            "   Calculates dimensions in inches and square feet from a text entry (e.g., \"5'2\\\" x 8'\").\n"
            "6. BULK Process Rug Sizes from File:\n"
            "   Processes a column of dimensions in an Excel/CSV file, adding calculated width, height, and area.\n"
            "7. Unit Converter:\n"
            "   Quickly converts between units like cm, m, ft, and inches.\n"
            "8. Match Image Links:\n"
            "   Matches image links from a separate file to a key column in an Excel/CSV file and adds them as new columns.\n"
            "---------------------------------\n"
            "Created by Hakan Akaslan"
        ),
    },
    "tr": {
        "Combined Utility Tool": "Birleşik Araç Aracı",
        "Welcome to the Combined Utility Tool!": "Birleşik Araç Aracına hoş geldiniz!",
        "File & Image Tools": "Dosya ve Görsel Araçları",
        "Data & Calculation": "Veri ve Hesaplama",
        "Code Generators": "Kod Üreteçleri",
        "Help & About": "Yardım ve Hakkında",
        "Language": "Dil",
        "English": "İngilizce",
        "Turkish": "Türkçe",
        "Language changed to {language}.": "Dil {language} olarak değiştirildi.",
        "1. Copy/Move Files by List": "1. Listeye Göre Dosya Kopyala/Taşı",
        "Source Folder:": "Kaynak Klasör:",
        "Target Folder:": "Hedef Klasör:",
        "Numbers File (List):": "Numara Dosyası (Liste):",
        "Browse...": "Gözat...",
        "Copy Files": "Dosyaları Kopyala",
        "Move Files": "Dosyaları Taşı",
        "Save Settings": "Ayarları Kaydet",
        "2. Convert HEIC to JPG": "2. HEIC'i JPG'ye Dönüştür",
        "Folder with HEIC files:": "HEIC dosyalarının olduğu klasör:",
        "Convert": "Dönüştür",
        "3. Batch Image Resizer": "3. Toplu Görsel Boyutlandırıcı",
        "Image Folder:": "Görsel Klasörü:",
        "Resize Mode:": "Yeniden Boyutlandırma Modu:",
        "By Width": "Genişliğe Göre",
        "By Percentage": "Yüzdeye Göre",
        "Max Width:": "Azami Genişlik:",
        "Percentage (%):": "Yüzde (%):",
        "JPEG Quality (1-95):": "JPEG Kalitesi (1-95):",
        "Resize & Compress": "Yeniden Boyutlandır ve Sıkıştır",
        "4. Format Numbers from File": "4. Dosyadan Numaraları Biçimlendir",
        "Excel/CSV/TXT File:": "Excel/CSV/TXT Dosyası:",
        "Format": "Biçimlendir",
        "5. Rug Size Calculator (Single)": "5. Halı Boyutu Hesaplayıcı (Tek)",
        "Dimension (e.g., 5'2\" x 8'):": "Ölçü (örn. 5'2\" x 8'):",
        "Calculate": "Hesapla",
        "6. BULK Process Rug Sizes from File": "6. Dosyadan Toplu Halı Ölçüsü İşle",
        "Excel/CSV File:": "Excel/CSV Dosyası:",
        "Column Name/Letter:": "Sütun Adı/Harf:",
        "Process File": "Dosyayı İşle",
        "7. Unit Converter": "7. Birim Dönüştürücü",
        "Conversion:": "Dönüşüm:",
        "182 cm to ft": "182 cm'yi ft'ye",
        "8. Match Image Links": "8. Görsel Bağlantılarını Eşleştir",
        "Source Excel/CSV File:": "Kaynak Excel/CSV Dosyası",
        "Image Links File (CSV):": "Görsel Bağlantı Dosyası (CSV):",
        "Key Column Name/Letter:": "Anahtar Sütun Adı/Harf:",
        "Match and Add Links": "Bağlantıları Eşleştir ve Ekle",
        "8. QR Code Generator": "8. QR Kod Oluşturucu",
        "Data/URL:": "Veri/URL:",
        "Output Type:": "Çıktı Türü:",
        "Standard PNG": "Standart PNG",
        "Dymo Label": "Dymo Etiketi",
        "Dymo Size:": "Dymo Boyutu",
        "Bottom Text:": "Alt Metin",
        "Filename:": "Dosya Adı",
        "Generate QR Code": "QR Kod Oluştur",
        "9. Barcode Generator": "9. Barkod Oluşturucu",
        "Data:": "Veri",
        "Format:": "Format",
        "Output Type:": "Çıktı Türü",
        "Generate Barcode": "Barkod Oluştur",
        "Check for Updates": "Güncellemeleri Kontrol Et",
        "Warning": "Uyarı",
        "Source and Target folders cannot be empty.": "Kaynak ve hedef klasörler boş olamaz.",
        "✅ Settings saved to settings.json": "✅ Ayarlar settings.json dosyasına kaydedildi",
        "Success": "Başarılı",
        "Folder settings have been saved.": "Klasör ayarları kaydedildi.",
        "Error": "Hata",
        "Please specify Source, Target, and Numbers File.": "Lütfen Kaynak, Hedef ve Numara Dosyasını belirtin.",
        "Please select a valid folder.": "Lütfen geçerli bir klasör seçin.",
        "Please select a valid image folder.": "Lütfen geçerli bir görsel klasörü seçin.",
        "Resize values and quality must be valid numbers.": "Yeniden boyutlandırma değerleri ve kalite geçerli sayılar olmalıdır.",
        "Please select a file.": "Lütfen bir dosya seçin.",
        "Please enter a dimension.": "Lütfen bir ölçü girin.",
        "Invalid Format": "Geçersiz Format",
        "W: {width} in | H: {height} in | Area: {area} sqft": "G: {width} in | Y: {height} in | Alan: {area} ft²",
        "Please select a file and specify a column.": "Lütfen bir dosya seçin ve bir sütun belirtin.",
        "Please fill in all file paths and the column name.": "Lütfen tüm dosya yollarını ve sütun adını doldurun.",
        "Data and filename are required.": "Veri ve dosya adı gereklidir.",
        "Error: {message}": "Hata: {message}",
        "ABOUT_CONTENT": (
            "Birleşik Araç Aracı - v{version}\n"
            "Bu uygulama yaygın dosya, görsel ve veri işleme görevlerini tek bir arayüzde toplar.\n"
            "--- ÖZELLİKLER ---\n"
            "1. Listeye Göre Dosya Kopyala/Taşı:\n"
            "   Excel veya metin dosyasındaki bir listeye göre görsel dosyaları bulur ve kopyalar/taşır.\n"
            "2. HEIC'i JPG'ye Dönüştür:\n"
            "   Apple'ın HEIC formatındaki görsellerini evrensel JPG formatına dönüştürür.\n"
            "3. Toplu Görsel Boyutlandırıcı:\n"
            "   Görselleri sabit bir genişliğe veya orijinal boyutların yüzdesine göre yeniden boyutlandırır.\n"
            "4. Dosyadan Numaraları Biçimlendir:\n"
            "   Bir dosyanın ilk sütunundaki öğeleri tek satırlık virgüllü listeye dönüştürür.\n"
            "5. Halı Boyutu Hesaplayıcı (Tek):\n"
            "   Metin girişinden inç ve metrekare hesaplar (örn. \"5'2\\\" x 8'\").\n"
            "6. Dosyadan Toplu Halı Ölçüsü İşle:\n"
            "   Excel/CSV dosyasındaki bir sütunu işleyip hesaplanan genişlik, yükseklik ve alan ekler.\n"
            "7. Birim Dönüştürücü:\n"
            "   cm, m, ft ve inç gibi birimler arasında hızlıca dönüştürme yapar.\n"
            "8. Görsel Bağlantılarını Eşleştir:\n"
            "   Ayrı bir dosyadaki görsel bağlantılarını Excel/CSV dosyasındaki anahtar sütuna eşleştirip yeni sütunlar olarak ekler.\n"
            "---------------------------------\n"
            "Geliştirici: Hakan Akaslan"
        ),
    },
}

DYMO_LABELS = {
    'Address (30252)': {'w_in': 3.5, 'h_in': 1.125},
    'Shipping (30256)': {'w_in': 4.0, 'h_in': 2.3125},
    'Small Multipurpose (30336)': {'w_in': 2.125, 'h_in': 1.0},
    'File Folder (30258)': {'w_in': 3.5, 'h_in': 0.5625},
}


class ToolApp(tk.Tk):
    """Main application window that builds the entire tkinter interface."""

    def __init__(self):
        super().__init__()

        self.settings = load_settings()
        self.language = self.settings.get("language", "en")
        if self.language not in TRANSLATIONS:
            self.language = "en"

        self.translatable_widgets = []
        self.translatable_tabs = []

        self.geometry("900x750")

        self.setup_styles()
        self.create_header()

        self.language_var = tk.StringVar(value=self.language)

        self.create_language_selector()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)

        self.create_file_image_tab()
        self.create_data_calc_tab()
        self.create_code_gen_tab()
        self.create_about_tab()

        self.log_area = ScrolledText(self, height=12)
        self.log_area.pack(pady=10, padx=10, fill="both", expand=True)
        self.log_area.config(
            state=tk.DISABLED,
            background="#0b1120",
            foreground="#f1f5f9",
            insertbackground="#f1f5f9",
            font=("Cascadia Code", 10),
            relief="flat",
            borderwidth=0,
        )

        self.refresh_translations()
        self.log(self.tr("Welcome to the Combined Utility Tool!"))

        self.run_in_thread(check_for_updates, self, self.log, __version__, silent=True)

    def setup_styles(self):
        """Configure a modern dark theme for the application widgets."""
        base_bg = "#0b1120"
        card_bg = "#111c2e"
        accent = "#38bdf8"
        accent_hover = "#0ea5e9"
        text_primary = "#f1f5f9"
        text_secondary = "#cbd5f5"
        text_muted = "#94a3b8"

        self.theme_colors = {
            "base_bg": base_bg,
            "card_bg": card_bg,
            "accent": accent,
            "accent_hover": accent_hover,
            "text_primary": text_primary,
            "text_secondary": text_secondary,
            "text_muted": text_muted,
        }

        self.configure(bg=base_bg)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=base_bg)
        style.configure("Header.TFrame", background=base_bg)
        style.configure("Card.TLabelframe", background=card_bg, borderwidth=0, padding=15)
        style.configure(
            "Card.TLabelframe.Label",
            background=card_bg,
            foreground=text_primary,
            font=("Segoe UI Semibold", 11),
        )
        style.configure(
            "TLabel",
            background=card_bg,
            foreground=text_secondary,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Primary.TLabel",
            background=base_bg,
            foreground=text_primary,
            font=("Segoe UI Semibold", 18),
        )
        style.configure(
            "Secondary.TLabel",
            background=base_bg,
            foreground=text_muted,
            font=("Segoe UI", 11),
        )
        style.configure(
            "TNotebook",
            background=base_bg,
            borderwidth=0,
            tabmargins=(4, 2, 4, 0),
        )
        style.configure(
            "TNotebook.Tab",
            background=card_bg,
            foreground=text_secondary,
            padding=(16, 8),
            font=("Segoe UI", 10),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", accent), ("active", accent_hover)],
            foreground=[("selected", base_bg), ("active", base_bg)],
        )
        style.configure(
            "TButton",
            background=accent,
            foreground=base_bg,
            font=("Segoe UI Semibold", 10),
            padding=(14, 6),
            borderwidth=0,
        )
        style.map(
            "TButton",
            background=[("active", accent_hover), ("disabled", "#1e293b")],
            foreground=[("disabled", text_muted)],
        )
        style.configure(
            "TEntry",
            fieldbackground="#111827",
            foreground=text_primary,
            insertcolor=text_primary,
            padding=8,
        )
        style.configure(
            "TLabelframe",
            background=card_bg,
            foreground=text_primary,
        )
        style.configure(
            "TRadiobutton",
            background=card_bg,
            foreground=text_primary,
            font=("Segoe UI", 10),
        )
        style.configure("Horizontal.TSeparator", background="#1f2937")

        self.option_add("*TCombobox*Listbox.font", ("Segoe UI", 10))
        self.option_add("*Font", ("Segoe UI", 10))
        self.option_add("*Foreground", text_primary)

    def create_header(self):
        """Create a simple branded header for the application."""
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x", padx=10, pady=(10, 0))

        title = ttk.Label(
            header,
            text=f"{self.tr('Combined Utility Tool')} v{__version__}",
            style="Primary.TLabel",
        )
        title.pack(anchor="w")
        self.header_title = title

        subtitle = ttk.Label(
            header,
            text=self.tr("Welcome to the Combined Utility Tool!"),
            style="Secondary.TLabel",
        )
        subtitle.pack(anchor="w", pady=(2, 0))
        self.header_subtitle = subtitle

    def tr(self, text_key):
        """Translate a text key according to the selected language."""
        return TRANSLATIONS.get(self.language, TRANSLATIONS["en"]).get(text_key, text_key)

    def register_widget(self, widget, text_key, attr="text"):
        """Register a widget for translation updates."""
        self.translatable_widgets.append((widget, attr, text_key))
        self._apply_translation(widget, attr, text_key)

    def register_tab(self, tab, text_key):
        """Register a notebook tab for translation updates."""
        self.translatable_tabs.append((tab, text_key))
        self.notebook.tab(tab, text=self.tr(text_key))

    def _apply_translation(self, widget, attr, text_key):
        try:
            widget.configure(**{attr: self.tr(text_key)})
        except tk.TclError:
            pass

    def refresh_translations(self):
        """Refresh UI texts according to the currently selected language."""
        self.title(f"{self.tr('Combined Utility Tool')} v{__version__}")
        if hasattr(self, "header_title"):
            self.header_title.config(text=f"{self.tr('Combined Utility Tool')} v{__version__}")
        if hasattr(self, "header_subtitle"):
            self.header_subtitle.config(text=self.tr("Welcome to the Combined Utility Tool!"))
        for widget, attr, text_key in self.translatable_widgets:
            self._apply_translation(widget, attr, text_key)
        for tab, text_key in self.translatable_tabs:
            self.notebook.tab(tab, text=self.tr(text_key))
        self.update_help_tab_content()

    def update_help_tab_content(self):
        if hasattr(self, "help_text_area"):
            self.help_text_area.config(state=tk.NORMAL)
            self.help_text_area.delete("1.0", tk.END)
            help_content = self.tr("ABOUT_CONTENT").format(version=__version__)
            self.help_text_area.insert(tk.END, help_content)
            self.help_text_area.config(state=tk.DISABLED)

    def change_language(self):
        """Handle changes coming from the language selector."""
        new_language = self.language_var.get()
        if new_language not in TRANSLATIONS or new_language == self.language:
            return
        self.language = new_language
        self.settings["language"] = self.language
        save_settings(self.settings)
        self.refresh_translations()
        self.log(self.tr("Language changed to {language}.").format(language=self.tr(self.language_name(self.language))))

    @staticmethod
    def language_name(code):
        return "English" if code == "en" else "Turkish"

    def create_language_selector(self):
        """Create the language selection controls placed at the top."""
        frame = ttk.LabelFrame(
            self,
            padding=(16, 12),
            text=self.tr("Language"),
            style="Card.TLabelframe",
        )
        frame.pack(fill="x", padx=10, pady=10)
        self.register_widget(frame, "Language")

        radio_en = ttk.Radiobutton(
            frame,
            text=self.tr("English"),
            value="en",
            variable=self.language_var,
            command=self.change_language,
        )
        radio_en.pack(side="left", padx=8)
        self.register_widget(radio_en, "English")

        radio_tr = ttk.Radiobutton(
            frame,
            text=self.tr("Turkish"),
            value="tr",
            variable=self.language_var,
            command=self.change_language,
        )
        radio_tr.pack(side="left", padx=8)
        self.register_widget(radio_tr, "Turkish")

    def log(self, message):
        if not isinstance(message, str):
            message = str(message)

        self.log_area.config(state=tk.NORMAL)

        if '\r' in message:
            self.log_area.delete("end-1l", "end")
            self.log_area.insert(tk.END, message.replace('\r', ''))
        else:
            self.log_area.insert(tk.END, message + "\n")

        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def run_in_thread(self, target_func, *args, **kwargs):
        thread = threading.Thread(target=target_func, args=args, kwargs=kwargs)
        thread.daemon = True
        thread.start()

    def task_completion_popup(self, title, message):
        """Shows a messagebox popup from the main thread."""
        self.after(0, messagebox.showinfo, self.tr(title), message)

    def create_file_image_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="")
        self.register_tab(tab, "File & Image Tools")

        file_ops_frame = ttk.LabelFrame(tab, text=self.tr("1. Copy/Move Files by List"))
        self.register_widget(file_ops_frame, "1. Copy/Move Files by List")
        file_ops_frame.pack(fill="x", padx=10, pady=10)

        self.source_folder = tk.StringVar(value=self.settings.get("source_folder", ""))
        self.target_folder = tk.StringVar(value=self.settings.get("target_folder", ""))
        self.numbers_file = tk.StringVar()

        src_label = ttk.Label(file_ops_frame, text=self.tr("Source Folder:"))
        src_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(src_label, "Source Folder:")
        ttk.Entry(file_ops_frame, textvariable=self.source_folder, width=60).grid(row=0, column=1, padx=5, pady=5)
        src_browse = ttk.Button(
            file_ops_frame,
            text=self.tr("Browse..."),
            command=lambda: self.source_folder.set(filedialog.askdirectory()),
        )
        src_browse.grid(row=0, column=2, padx=5, pady=5)
        self.register_widget(src_browse, "Browse...")

        tgt_label = ttk.Label(file_ops_frame, text=self.tr("Target Folder:"))
        tgt_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(tgt_label, "Target Folder:")
        ttk.Entry(file_ops_frame, textvariable=self.target_folder, width=60).grid(row=1, column=1, padx=5, pady=5)
        tgt_browse = ttk.Button(
            file_ops_frame,
            text=self.tr("Browse..."),
            command=lambda: self.target_folder.set(filedialog.askdirectory()),
        )
        tgt_browse.grid(row=1, column=2, padx=5, pady=5)
        self.register_widget(tgt_browse, "Browse...")

        numbers_label = ttk.Label(file_ops_frame, text=self.tr("Numbers File (List):"))
        numbers_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(numbers_label, "Numbers File (List):")
        ttk.Entry(file_ops_frame, textvariable=self.numbers_file, width=60).grid(row=2, column=1, padx=5, pady=5)
        numbers_browse = ttk.Button(
            file_ops_frame,
            text=self.tr("Browse..."),
            command=lambda: self.numbers_file.set(filedialog.askopenfilename()),
        )
        numbers_browse.grid(row=2, column=2, padx=5, pady=5)
        self.register_widget(numbers_browse, "Browse...")

        btn_frame = ttk.Frame(file_ops_frame)
        btn_frame.grid(row=3, column=1, pady=10)

        copy_btn = ttk.Button(btn_frame, text=self.tr("Copy Files"), command=lambda: self.start_process_files("copy"))
        copy_btn.pack(side="left", padx=5)
        self.register_widget(copy_btn, "Copy Files")

        move_btn = ttk.Button(btn_frame, text=self.tr("Move Files"), command=lambda: self.start_process_files("move"))
        move_btn.pack(side="left", padx=5)
        self.register_widget(move_btn, "Move Files")

        save_btn = ttk.Button(btn_frame, text=self.tr("Save Settings"), command=self.save_folder_settings)
        save_btn.pack(side="left", padx=5)
        self.register_widget(save_btn, "Save Settings")

        heic_frame = ttk.LabelFrame(tab, text=self.tr("2. Convert HEIC to JPG"))
        self.register_widget(heic_frame, "2. Convert HEIC to JPG")
        heic_frame.pack(fill="x", padx=10, pady=10)

        self.heic_folder = tk.StringVar()
        heic_label = ttk.Label(heic_frame, text=self.tr("Folder with HEIC files:"))
        heic_label.pack(side="left", padx=5, pady=5)
        self.register_widget(heic_label, "Folder with HEIC files:")
        ttk.Entry(heic_frame, textvariable=self.heic_folder, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        heic_browse = ttk.Button(heic_frame, text=self.tr("Browse..."), command=lambda: self.heic_folder.set(filedialog.askdirectory()))
        heic_browse.pack(side="left", padx=5, pady=5)
        self.register_widget(heic_browse, "Browse...")
        heic_convert = ttk.Button(heic_frame, text=self.tr("Convert"), command=self.start_heic_conversion)
        heic_convert.pack(side="left", padx=5, pady=5)
        self.register_widget(heic_convert, "Convert")

        resize_frame = ttk.LabelFrame(tab, text=self.tr("3. Batch Image Resizer"))
        self.register_widget(resize_frame, "3. Batch Image Resizer")
        resize_frame.pack(fill="x", padx=10, pady=10)

        self.resize_folder = tk.StringVar()
        self.quality = tk.StringVar(value="75")
        self.resize_mode = tk.StringVar(value="width")
        self.max_width = tk.StringVar(value="1920")
        self.resize_percentage = tk.StringVar(value="50")

        resize_folder_label = ttk.Label(resize_frame, text=self.tr("Image Folder:"))
        resize_folder_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(resize_folder_label, "Image Folder:")
        ttk.Entry(resize_frame, textvariable=self.resize_folder, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        resize_browse = ttk.Button(resize_frame, text=self.tr("Browse..."), command=lambda: self.resize_folder.set(filedialog.askdirectory()))
        resize_browse.grid(row=0, column=4, padx=5, pady=5)
        self.register_widget(resize_browse, "Browse...")

        resize_mode_label = ttk.Label(resize_frame, text=self.tr("Resize Mode:"))
        resize_mode_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(resize_mode_label, "Resize Mode:")

        radio_frame = ttk.Frame(resize_frame)
        radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

        radio_width = ttk.Radiobutton(radio_frame, text=self.tr("By Width"), variable=self.resize_mode, value="width", command=self.toggle_resize_mode)
        radio_width.pack(side="left")
        self.register_widget(radio_width, "By Width")

        radio_percentage = ttk.Radiobutton(radio_frame, text=self.tr("By Percentage"), variable=self.resize_mode, value="percentage", command=self.toggle_resize_mode)
        radio_percentage.pack(side="left", padx=10)
        self.register_widget(radio_percentage, "By Percentage")

        max_width_label = ttk.Label(resize_frame, text=self.tr("Max Width:"))
        max_width_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(max_width_label, "Max Width:")
        self.width_entry = ttk.Entry(resize_frame, textvariable=self.max_width, width=10)
        self.width_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        percentage_label = ttk.Label(resize_frame, text=self.tr("Percentage (%):"))
        percentage_label.grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.register_widget(percentage_label, "Percentage (%):")
        self.percentage_entry = ttk.Entry(resize_frame, textvariable=self.resize_percentage, width=10)
        self.percentage_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        jpeg_quality_label = ttk.Label(resize_frame, text=self.tr("JPEG Quality (1-95):"))
        jpeg_quality_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(jpeg_quality_label, "JPEG Quality (1-95):")
        ttk.Entry(resize_frame, textvariable=self.quality, width=10).grid(row=3, column=1, padx=5, pady=5, sticky="w")

        resize_button = ttk.Button(resize_frame, text=self.tr("Resize & Compress"), command=self.start_resize_task)
        resize_button.grid(row=4, column=1, columnspan=2, pady=10)
        self.register_widget(resize_button, "Resize & Compress")

        self.toggle_resize_mode()

    def toggle_resize_mode(self):
        if self.resize_mode.get() == "width":
            self.width_entry.config(state="normal")
            self.percentage_entry.config(state="disabled")
        else:
            self.width_entry.config(state="disabled")
            self.percentage_entry.config(state="normal")

    def create_data_calc_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="")
        self.register_tab(tab, "Data & Calculation")

        format_frame = ttk.LabelFrame(tab, text=self.tr("4. Format Numbers from File"))
        self.register_widget(format_frame, "4. Format Numbers from File")
        format_frame.pack(fill="x", padx=10, pady=10)

        self.format_file = tk.StringVar()
        format_label = ttk.Label(format_frame, text=self.tr("Excel/CSV/TXT File:"))
        format_label.pack(side="left", padx=5, pady=5)
        self.register_widget(format_label, "Excel/CSV/TXT File:")
        ttk.Entry(format_frame, textvariable=self.format_file, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        format_browse = ttk.Button(format_frame, text=self.tr("Browse..."), command=lambda: self.format_file.set(filedialog.askopenfilename()))
        format_browse.pack(side="left", padx=5, pady=5)
        self.register_widget(format_browse, "Browse...")
        format_button = ttk.Button(format_frame, text=self.tr("Format"), command=self.start_format_numbers)
        format_button.pack(side="left", padx=5, pady=5)
        self.register_widget(format_button, "Format")

        single_rug_frame = ttk.LabelFrame(tab, text=self.tr("5. Rug Size Calculator (Single)"))
        self.register_widget(single_rug_frame, "5. Rug Size Calculator (Single)")
        single_rug_frame.pack(fill="x", padx=10, pady=10)

        self.rug_dim_input = tk.StringVar()
        self.rug_result_label = tk.StringVar()

        rug_label = ttk.Label(single_rug_frame, text=self.tr("Dimension (e.g., 5'2\" x 8'):"))
        rug_label.pack(side="left", padx=5, pady=5)
        self.register_widget(rug_label, "Dimension (e.g., 5'2\" x 8'):")
        ttk.Entry(single_rug_frame, textvariable=self.rug_dim_input, width=20).pack(side="left", padx=5, pady=5)
        rug_button = ttk.Button(single_rug_frame, text=self.tr("Calculate"), command=self.calculate_single_rug)
        rug_button.pack(side="left", padx=5, pady=5)
        self.register_widget(rug_button, "Calculate")
        ttk.Label(single_rug_frame, textvariable=self.rug_result_label, font=("Helvetica", 10, "bold")).pack(side="left", padx=15, pady=5)

        bulk_rug_frame = ttk.LabelFrame(tab, text=self.tr("6. BULK Process Rug Sizes from File"))
        self.register_widget(bulk_rug_frame, "6. BULK Process Rug Sizes from File")
        bulk_rug_frame.pack(fill="x", padx=10, pady=10)

        self.bulk_rug_file = tk.StringVar()
        self.bulk_rug_col = tk.StringVar(value="Size")

        bulk_file_label = ttk.Label(bulk_rug_frame, text=self.tr("Excel/CSV File:"))
        bulk_file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(bulk_file_label, "Excel/CSV File:")
        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        bulk_browse = ttk.Button(bulk_rug_frame, text=self.tr("Browse..."), command=lambda: self.bulk_rug_file.set(filedialog.askopenfilename()))
        bulk_browse.grid(row=0, column=2, padx=5, pady=5)
        self.register_widget(bulk_browse, "Browse...")

        bulk_col_label = ttk.Label(bulk_rug_frame, text=self.tr("Column Name/Letter:"))
        bulk_col_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(bulk_col_label, "Column Name/Letter:")
        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_col, width=20).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        bulk_process = ttk.Button(bulk_rug_frame, text=self.tr("Process File"), command=self.start_bulk_rug_sizer)
        bulk_process.grid(row=1, column=2, padx=5, pady=5)
        self.register_widget(bulk_process, "Process File")

        unit_frame = ttk.LabelFrame(tab, text=self.tr("7. Unit Converter"))
        self.register_widget(unit_frame, "7. Unit Converter")
        unit_frame.pack(fill="x", padx=10, pady=10)

        self.unit_input = tk.StringVar(value=self.tr("182 cm to ft"))
        self.unit_result_label = tk.StringVar()

        conversion_label = ttk.Label(unit_frame, text=self.tr("Conversion:"))
        conversion_label.pack(side="left", padx=5, pady=5)
        self.register_widget(conversion_label, "Conversion:")
        ttk.Entry(unit_frame, textvariable=self.unit_input, width=20).pack(side="left", padx=5, pady=5)
        convert_button = ttk.Button(unit_frame, text=self.tr("Convert"), command=self.convert_units)
        convert_button.pack(side="left", padx=5, pady=5)
        self.register_widget(convert_button, "Convert")
        ttk.Label(unit_frame, textvariable=self.unit_result_label, font=("Helvetica", 10, "bold")).pack(side="left", padx=15, pady=5)

        image_link_frame = ttk.LabelFrame(tab, text=self.tr("8. Match Image Links"))
        self.register_widget(image_link_frame, "8. Match Image Links")
        image_link_frame.pack(fill="x", padx=10, pady=10)

        self.input_excel_file = tk.StringVar()
        self.image_links_file = tk.StringVar(value="image link shopify.csv")
        self.key_column = tk.StringVar(value="A")

        source_file_label = ttk.Label(image_link_frame, text=self.tr("Source Excel/CSV File:"))
        source_file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(source_file_label, "Source Excel/CSV File:")
        ttk.Entry(image_link_frame, textvariable=self.input_excel_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        source_browse = ttk.Button(
            image_link_frame,
            text=self.tr("Browse..."),
            command=lambda: self.input_excel_file.set(
                filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
            ),
        )
        source_browse.grid(row=0, column=2, padx=5, pady=5)
        self.register_widget(source_browse, "Browse...")

        image_links_label = ttk.Label(image_link_frame, text=self.tr("Image Links File (CSV):"))
        image_links_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(image_links_label, "Image Links File (CSV):")
        ttk.Entry(image_link_frame, textvariable=self.image_links_file, width=50).grid(row=1, column=1, padx=5, pady=5)
        image_links_browse = ttk.Button(
            image_link_frame,
            text=self.tr("Browse..."),
            command=lambda: self.image_links_file.set(filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])),
        )
        image_links_browse.grid(row=1, column=2, padx=5, pady=5)
        self.register_widget(image_links_browse, "Browse...")

        key_column_label = ttk.Label(image_link_frame, text=self.tr("Key Column Name/Letter:"))
        key_column_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.register_widget(key_column_label, "Key Column Name/Letter:")
        ttk.Entry(image_link_frame, textvariable=self.key_column, width=10).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        match_button = ttk.Button(image_link_frame, text=self.tr("Match and Add Links"), command=self.start_add_image_links)
        match_button.grid(row=3, column=1, pady=10)
        self.register_widget(match_button, "Match and Add Links")

    def create_code_gen_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="")
        self.register_tab(tab, "Code Generators")

        def toggle_dymo_options(output_var, combobox, entry):
            if output_var.get() == "Dymo":
                combobox.config(state="readonly")
                entry.config(state="normal")
            else:
                combobox.config(state="disabled")
                entry.config(state="disabled")

        qr_frame = ttk.LabelFrame(tab, text=self.tr("8. QR Code Generator"))
        self.register_widget(qr_frame, "8. QR Code Generator")
        qr_frame.pack(fill="x", padx=10, pady=10)

        self.qr_data = tk.StringVar()
        self.qr_filename = tk.StringVar(value="qrcode.png")
        self.qr_output_type = tk.StringVar(value="PNG")
        self.qr_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.qr_bottom_text = tk.StringVar()

        qr_data_label = ttk.Label(qr_frame, text=self.tr("Data/URL:"))
        qr_data_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(qr_data_label, "Data/URL:")
        ttk.Entry(qr_frame, textvariable=self.qr_data, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5)

        qr_output_label = ttk.Label(qr_frame, text=self.tr("Output Type:"))
        qr_output_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(qr_output_label, "Output Type:")

        qr_radio_frame = ttk.Frame(qr_frame)
        qr_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

        qr_dymo_combo = ttk.Combobox(qr_frame, textvariable=self.qr_dymo_size, values=list(DYMO_LABELS.keys()), state="disabled", width=30)
        qr_bottom_entry = ttk.Entry(qr_frame, textvariable=self.qr_bottom_text, state="disabled", width=32)

        qr_png_radio = ttk.Radiobutton(
            qr_radio_frame,
            text=self.tr("Standard PNG"),
            variable=self.qr_output_type,
            value="PNG",
            command=lambda: toggle_dymo_options(self.qr_output_type, qr_dymo_combo, qr_bottom_entry),
        )
        qr_png_radio.pack(side="left", padx=5)
        self.register_widget(qr_png_radio, "Standard PNG")

        qr_dymo_radio = ttk.Radiobutton(
            qr_radio_frame,
            text=self.tr("Dymo Label"),
            variable=self.qr_output_type,
            value="Dymo",
            command=lambda: toggle_dymo_options(self.qr_output_type, qr_dymo_combo, qr_bottom_entry),
        )
        qr_dymo_radio.pack(side="left", padx=5)
        self.register_widget(qr_dymo_radio, "Dymo Label")

        qr_dymo_label = ttk.Label(qr_frame, text=self.tr("Dymo Size:"))
        qr_dymo_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(qr_dymo_label, "Dymo Size:")
        qr_dymo_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        qr_bottom_label = ttk.Label(qr_frame, text=self.tr("Bottom Text:"))
        qr_bottom_label.grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.register_widget(qr_bottom_label, "Bottom Text:")
        qr_bottom_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        qr_filename_label = ttk.Label(qr_frame, text=self.tr("Filename:"))
        qr_filename_label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(qr_filename_label, "Filename:")
        ttk.Entry(qr_frame, textvariable=self.qr_filename, width=60).grid(row=3, column=1, columnspan=3, padx=5, pady=5)

        qr_button = ttk.Button(qr_frame, text=self.tr("Generate QR Code"), command=self.start_generate_qr)
        qr_button.grid(row=4, column=1, columnspan=2, pady=10)
        self.register_widget(qr_button, "Generate QR Code")

        bc_frame = ttk.LabelFrame(tab, text=self.tr("9. Barcode Generator"))
        self.register_widget(bc_frame, "9. Barcode Generator")
        bc_frame.pack(fill="x", padx=10, pady=10)

        self.bc_data = tk.StringVar()
        self.bc_filename = tk.StringVar(value="barcode.png")
        self.bc_type = tk.StringVar(value='code128')
        self.bc_output_type = tk.StringVar(value="PNG")
        self.bc_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.bc_bottom_text = tk.StringVar()

        bc_data_label = ttk.Label(bc_frame, text=self.tr("Data:"))
        bc_data_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(bc_data_label, "Data:")
        ttk.Entry(bc_frame, textvariable=self.bc_data, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        bc_format_label = ttk.Label(bc_frame, text=self.tr("Format:"))
        bc_format_label.grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.register_widget(bc_format_label, "Format:")
        ttk.Combobox(bc_frame, textvariable=self.bc_type, values=['code39', 'code128', 'ean13', 'upca'], state="readonly", width=15).grid(row=0, column=3, padx=5, pady=5, sticky="w")

        bc_output_label = ttk.Label(bc_frame, text=self.tr("Output Type:"))
        bc_output_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(bc_output_label, "Output Type:")

        bc_radio_frame = ttk.Frame(bc_frame)
        bc_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

        bc_dymo_combo = ttk.Combobox(bc_frame, textvariable=self.bc_dymo_size, values=list(DYMO_LABELS.keys()), state="disabled", width=30)
        bc_bottom_entry = ttk.Entry(bc_frame, textvariable=self.bc_bottom_text, state="disabled", width=32)

        bc_png_radio = ttk.Radiobutton(
            bc_radio_frame,
            text=self.tr("Standard PNG"),
            variable=self.bc_output_type,
            value="PNG",
            command=lambda: toggle_dymo_options(self.bc_output_type, bc_dymo_combo, bc_bottom_entry),
        )
        bc_png_radio.pack(side="left", padx=5)
        self.register_widget(bc_png_radio, "Standard PNG")

        bc_dymo_radio = ttk.Radiobutton(
            bc_radio_frame,
            text=self.tr("Dymo Label"),
            variable=self.bc_output_type,
            value="Dymo",
            command=lambda: toggle_dymo_options(self.bc_output_type, bc_dymo_combo, bc_bottom_entry),
        )
        bc_dymo_radio.pack(side="left", padx=5)
        self.register_widget(bc_dymo_radio, "Dymo Label")

        bc_dymo_label = ttk.Label(bc_frame, text=self.tr("Dymo Size:"))
        bc_dymo_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(bc_dymo_label, "Dymo Size:")
        bc_dymo_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        bc_bottom_label = ttk.Label(bc_frame, text=self.tr("Bottom Text:"))
        bc_bottom_label.grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.register_widget(bc_bottom_label, "Bottom Text:")
        bc_bottom_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        bc_filename_label = ttk.Label(bc_frame, text=self.tr("Filename:"))
        bc_filename_label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.register_widget(bc_filename_label, "Filename:")
        ttk.Entry(bc_frame, textvariable=self.bc_filename, width=60).grid(row=3, column=1, columnspan=3, padx=5, pady=5)

        bc_button = ttk.Button(bc_frame, text=self.tr("Generate Barcode"), command=self.start_generate_barcode)
        bc_button.grid(row=4, column=1, columnspan=2, pady=10)
        self.register_widget(bc_button, "Generate Barcode")

    def create_about_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="")
        self.register_tab(tab, "Help & About")

        top_frame = ttk.Frame(tab)
        top_frame.pack(fill="x", padx=10, pady=5)

        update_button = ttk.Button(top_frame, text=self.tr("Check for Updates"), command=lambda: self.run_in_thread(check_for_updates, self, self.log, __version__, silent=False))
        update_button.pack(side="left")
        self.register_widget(update_button, "Check for Updates")

        self.help_text_area = ScrolledText(tab, wrap=tk.WORD, padx=10, pady=10, font=("Helvetica", 10))
        self.help_text_area.configure(
            background=self.theme_colors["card_bg"],
            foreground=self.theme_colors["text_primary"],
            insertbackground=self.theme_colors["text_primary"],
            highlightthickness=0,
            borderwidth=0,
        )
        try:
            self.help_text_area.configure(
                disabledforeground=self.theme_colors["text_primary"],
                disabledbackground=self.theme_colors["card_bg"],
            )
        except tk.TclError:
            # Some Tk builds do not support disabled foreground/background options.
            pass
        self.help_text_area.pack(fill="both", expand=True)
        self.update_help_tab_content()

    def save_folder_settings(self):
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        if not src or not tgt:
            messagebox.showwarning(self.tr("Warning"), self.tr("Source and Target folders cannot be empty."))
            return
        self.settings['source_folder'] = src
        self.settings['target_folder'] = tgt
        save_settings(self.settings)
        self.log(self.tr("✅ Settings saved to settings.json"))
        messagebox.showinfo(self.tr("Success"), self.tr("Folder settings have been saved."))

    def start_process_files(self, action):
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        nums_f = self.numbers_file.get()
        if not all([src, tgt, nums_f]):
            messagebox.showerror(self.tr("Error"), self.tr("Please specify Source, Target, and Numbers File."))
            return
        self.run_in_thread(backend.process_files_task, src, tgt, nums_f, action, self.log, self.task_completion_popup)

    def start_heic_conversion(self):
        folder = self.heic_folder.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror(self.tr("Error"), self.tr("Please select a valid folder."))
            return
        self.run_in_thread(backend.convert_heic_task, folder, self.log, self.task_completion_popup)

    def start_resize_task(self):
        src_folder = self.resize_folder.get()
        if not src_folder or not os.path.isdir(src_folder):
            messagebox.showerror(self.tr("Error"), self.tr("Please select a valid image folder."))
            return
        mode = self.resize_mode.get()
        try:
            value = int(self.max_width.get()) if mode == 'width' else int(self.resize_percentage.get())
            quality = int(self.quality.get())
            if not (value > 0 and 1 <= quality <= 95):
                raise ValueError
        except ValueError:
            messagebox.showerror(self.tr("Error"), self.tr("Resize values and quality must be valid numbers."))
            return
        self.run_in_thread(backend.resize_images_task, src_folder, mode, value, quality, self.log, self.task_completion_popup)

    def start_format_numbers(self):
        file_path = self.format_file.get()
        if not file_path:
            messagebox.showerror(self.tr("Error"), self.tr("Please select a file."))
            return
        err, success_msg = backend.format_numbers_task(file_path)
        if err:
            self.log(self.tr("Error: {message}").format(message=err))
            messagebox.showerror(self.tr("Error"), err)
        else:
            self.log(success_msg)
            messagebox.showinfo(self.tr("Success"), success_msg)

    def calculate_single_rug(self):
        dim_str = self.rug_dim_input.get()
        if not dim_str:
            self.rug_result_label.set(self.tr("Please enter a dimension."))
            return
        w, h = backend.size_to_inches_wh(dim_str)
        s = backend.calculate_sqft(dim_str)
        if w is not None:
            self.rug_result_label.set(self.tr("W: {width} in | H: {height} in | Area: {area} sqft").format(width=w, height=h, area=s))
        else:
            self.rug_result_label.set(self.tr("Invalid Format"))

    def start_bulk_rug_sizer(self):
        path = self.bulk_rug_file.get()
        col = self.bulk_rug_col.get()
        if not path or not col:
            messagebox.showerror(self.tr("Error"), self.tr("Please select a file and specify a column."))
            return
        self.run_in_thread(backend.bulk_rug_sizer_task, path, col, self.log, self.task_completion_popup)

    def convert_units(self):
        input_str = self.unit_input.get()
        result_str = backend.convert_units_logic(input_str)
        self.unit_result_label.set(result_str)

    def start_add_image_links(self):
        input_path = self.input_excel_file.get()
        links_path = self.image_links_file.get()
        key_col = self.key_column.get()
        if not all([input_path, links_path, key_col]):
            messagebox.showerror(self.tr("Error"), self.tr("Please fill in all file paths and the column name."))
            return
        self.run_in_thread(
            backend.add_image_links_task,
            input_path,
            links_path,
            key_col,
            self.log,
            self.task_completion_popup,
        )

    def start_generate_qr(self):
        data = self.qr_data.get()
        fname = self.qr_filename.get()
        if not data or not fname:
            messagebox.showerror(self.tr("Error"), self.tr("Data and filename are required."))
            return
        dymo_info = DYMO_LABELS[self.qr_dymo_size.get()] if self.qr_output_type.get() == "Dymo" else None
        log_msg, success_msg = backend.generate_qr_task(data, fname, self.qr_output_type.get(), dymo_info, self.qr_bottom_text.get())
        self.log(log_msg)
        if success_msg:
            self.task_completion_popup("Success", success_msg)
        else:
            messagebox.showerror(self.tr("Error"), log_msg)

    def start_generate_barcode(self):
        data = self.bc_data.get()
        fname = self.bc_filename.get()
        if not data or not fname:
            messagebox.showerror(self.tr("Error"), self.tr("Data and filename are required."))
            return
        dymo_info = DYMO_LABELS[self.bc_dymo_size.get()] if self.bc_output_type.get() == "Dymo" else None
        log_msg, success_msg = backend.generate_barcode_task(
            data,
            fname,
            self.bc_type.get(),
            self.bc_output_type.get(),
            dymo_info,
            self.bc_bottom_text.get() or data,
        )
        self.log(log_msg)
        if success_msg:
            self.task_completion_popup("Success", success_msg)
        else:
            messagebox.showerror(self.tr("Error"), log_msg)

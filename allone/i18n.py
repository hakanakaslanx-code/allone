"""Translation helpers and dictionaries for the Combined Utility Tool UI."""
from __future__ import annotations

LANGUAGE_DISPLAY = {"en": "English", "tr": "Türkçe"}

TEXTS = {
    "en": {
        "app_title": "Combined Utility Tool v{version}",
        "welcome_log": "Welcome to the Combined Utility Tool!",
        "language_label": "Language:",
        "language_changed_log": "Language switched to {language}.",
        "tab_file_image": "File & Image Tools",
        "tab_data_calc": "Data & Calculation",
        "tab_code_gen": "Code Generators",
        "tab_about": "Help & About",
        "section_copy_move": "1. Copy/Move Files by List",
        "section_heic": "2. Convert HEIC to JPG",
        "section_resize": "3. Batch Image Resizer",
        "section_format_numbers": "4. Format Numbers from File",
        "section_rug_single": "5. Rug Size Calculator (Single)",
        "section_rug_bulk": "6. BULK Process Rug Sizes from File",
        "section_unit_converter": "7. Unit Converter",
        "section_match_links": "8. Match Image Links",
        "section_qr": "9. QR Code Generator",
        "section_barcode": "10. Barcode Generator",
        "label_source_folder": "Source Folder:",
        "label_target_folder": "Target Folder:",
        "label_numbers_file": "Numbers File (List):",
        "btn_browse": "Browse...",
        "btn_copy_files": "Copy Files",
        "btn_move_files": "Move Files",
        "btn_save_settings": "Save Settings",
        "label_heic_folder": "Folder with HEIC files:",
        "btn_convert": "Convert",
        "label_image_folder": "Image Folder:",
        "label_resize_mode": "Resize Mode:",
        "radio_by_width": "By Width",
        "radio_by_percentage": "By Percentage",
        "label_max_width": "Max Width:",
        "label_percentage": "Percentage (%):",
        "label_quality": "JPEG Quality (1-95):",
        "btn_resize": "Resize & Compress",
        "label_format_file": "Excel/CSV/TXT File:",
        "btn_format": "Format",
        "label_rug_dimension": "Dimension (e.g., 5'2\" x 8'):",
        "btn_calculate": "Calculate",
        "rug_enter_dimension": "Please enter a dimension.",
        "rug_invalid": "Invalid Format",
        "rug_result": "W: {width} in | H: {height} in | Area: {area} sqft",
        "label_bulk_file": "Excel/CSV File:",
        "label_bulk_column": "Column Name/Letter:",
        "btn_process_file": "Process File",
        "label_conversion": "Conversion:",
        "btn_convert_units": "Convert",
        "label_source_excel": "Source Excel/CSV File:",
        "label_image_links": "Image Links File (CSV):",
        "label_key_column": "Key Column Name/Letter:",
        "btn_match_links": "Match and Add Links",
        "label_data_url": "Data/URL:",
        "label_output_type": "Output Type:",
        "radio_png": "Standard PNG",
        "radio_dymo": "Dymo Label",
        "label_dymo_size": "Dymo Size:",
        "label_bottom_text": "Bottom Text:",
        "label_filename": "Filename:",
        "btn_generate_qr": "Generate QR Code",
        "label_barcode_data": "Data:",
        "label_format": "Format:",
        "btn_generate_barcode": "Generate Barcode",
        "btn_check_updates": "Check for Updates",
        "help_content": (
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
            "9. QR Code Generator:\n"
            "   Creates QR codes as PNG images or printable Dymo labels with optional bottom text.\n"
            "10. Barcode Generator:\n"
            "   Produces barcode images in popular formats or ready-to-print Dymo layouts.\n"
            "--- SUGGESTED ADDITIONS ---\n"
            "• Shopify product data cleaner to normalize titles, tags, and prices before upload.\n"
            "• Image background removal helper that links to local CLI tools or web APIs.\n"
            "• Bulk CSV quality checker to flag missing media, invalid UPCs, or duplicate handles.\n"
            "---------------------------------\n"
            "Created by Hakan Akaslan\n"
        ),
        "warning_title": "Warning",
        "settings_empty": "Source and Target folders cannot be empty.",
        "log_settings_saved": "✅ Settings saved to settings.json",
        "success_title": "Success",
        "settings_saved_popup": "Folder settings have been saved.",
        "error_title": "Error",
        "missing_required_fields": "Please specify Source, Target, and Numbers File.",
        "invalid_folder": "Please select a valid folder.",
        "resize_value_error": "Resize values and quality must be valid numbers.",
        "select_file": "Please select a file.",
        "missing_file_and_column": "Please select a file and specify a column.",
        "missing_inputs": "Please fill in all file paths and the column name.",
        "data_filename_required": "Data and filename are required.",
        "complete_title": "Complete",
    },
    "tr": {
        "app_title": "Birleşik Yardımcı Araç v{version}",
        "welcome_log": "Combined Utility Tool'a hoş geldiniz!",
        "language_label": "Dil:",
        "language_changed_log": "Dil {language} olarak değiştirildi.",
        "tab_file_image": "Dosya ve Görsel Araçları",
        "tab_data_calc": "Veri ve Hesaplama",
        "tab_code_gen": "Kod Üreticileri",
        "tab_about": "Yardım ve Hakkında",
        "section_copy_move": "1. Listeye Göre Dosya Kopyala/Taşı",
        "section_heic": "2. HEIC'i JPG'ye Dönüştür",
        "section_resize": "3. Toplu Görsel Yeniden Boyutlandırıcı",
        "section_format_numbers": "4. Dosyadan Sayıları Biçimlendir",
        "section_rug_single": "5. Halı Ölçü Hesaplayıcı (Tek)",
        "section_rug_bulk": "6. Dosyadan Toplu Halı Ölçüleri",
        "section_unit_converter": "7. Birim Dönüştürücü",
        "section_match_links": "8. Görsel Bağlantılarını Eşleştir",
        "section_qr": "9. QR Kod Üreticisi",
        "section_barcode": "10. Barkod Üreticisi",
        "label_source_folder": "Kaynak Klasör:",
        "label_target_folder": "Hedef Klasör:",
        "label_numbers_file": "Numara Dosyası (Liste):",
        "btn_browse": "Gözat...",
        "btn_copy_files": "Dosyaları Kopyala",
        "btn_move_files": "Dosyaları Taşı",
        "btn_save_settings": "Ayarları Kaydet",
        "label_heic_folder": "HEIC dosyalarının olduğu klasör:",
        "btn_convert": "Dönüştür",
        "label_image_folder": "Görsel Klasörü:",
        "label_resize_mode": "Boyutlandırma Modu:",
        "radio_by_width": "Genişliğe Göre",
        "radio_by_percentage": "Yüzdeye Göre",
        "label_max_width": "Azami Genişlik:",
        "label_percentage": "Yüzde (%):",
        "label_quality": "JPEG Kalitesi (1-95):",
        "btn_resize": "Yeniden Boyutlandır ve Sıkıştır",
        "label_format_file": "Excel/CSV/TXT Dosyası:",
        "btn_format": "Biçimlendir",
        "label_rug_dimension": "Ölçü (örn. 5'2\" x 8'):",
        "btn_calculate": "Hesapla",
        "rug_enter_dimension": "Lütfen bir ölçü girin.",
        "rug_invalid": "Geçersiz Format",
        "rug_result": "G: {width} inç | Y: {height} inç | Alan: {area} ft²",
        "label_bulk_file": "Excel/CSV Dosyası:",
        "label_bulk_column": "Sütun Adı/Harf:",
        "btn_process_file": "Dosyayı İşle",
        "label_conversion": "Dönüşüm:",
        "btn_convert_units": "Dönüştür",
        "label_source_excel": "Kaynak Excel/CSV Dosyası:",
        "label_image_links": "Görsel Bağlantı Dosyası (CSV):",
        "label_key_column": "Anahtar Sütun Adı/Harf:",
        "btn_match_links": "Bağlantıları Eşleştir ve Ekle",
        "label_data_url": "Veri/URL:",
        "label_output_type": "Çıktı Türü:",
        "radio_png": "Standart PNG",
        "radio_dymo": "Dymo Etiketi",
        "label_dymo_size": "Dymo Boyutu:",
        "label_bottom_text": "Alt Metin:",
        "label_filename": "Dosya Adı:",
        "btn_generate_qr": "QR Kod Oluştur",
        "label_barcode_data": "Veri:",
        "label_format": "Format:",
        "btn_generate_barcode": "Barkod Oluştur",
        "btn_check_updates": "Güncellemeleri Kontrol Et",
        "help_content": (
            "Birleşik Yardımcı Araç - v{version}\n"
            "Bu uygulama sık kullanılan dosya, görsel ve veri işlemlerini tek bir arayüzde toplar.\n"
            "--- ÖZELLİKLER ---\n"
            "1. Listeye Göre Dosya Kopyala/Taşı:\n"
            "   Excel veya metin dosyasındaki listeye göre görselleri bulur ve kopyalar/taşır.\n"
            "2. HEIC'i JPG'ye Dönüştür:\n"
            "   Apple'ın HEIC formatındaki görsellerini evrensel JPG biçimine çevirir.\n"
            "3. Toplu Görsel Yeniden Boyutlandırıcı:\n"
            "   Görselleri sabit genişliğe veya orijinal boyutun yüzdesine göre yeniden boyutlandırır.\n"
            "4. Dosyadan Sayıları Biçimlendir:\n"
            "   Dosyanın ilk sütunundaki değerleri tek satırlık virgüllü listeye dönüştürür.\n"
            "5. Halı Ölçü Hesaplayıcı (Tek):\n"
            "   Metin girişinden (örn. \"5'2\\\" x 8'\") inç ve metrekare karşılığını hesaplar.\n"
            "6. Dosyadan Toplu Halı Ölçüleri:\n"
            "   Excel/CSV dosyasındaki ölçü sütununu işleyip genişlik, yükseklik ve alan ekler.\n"
            "7. Birim Dönüştürücü:\n"
            "   cm, m, ft ve inç gibi birimler arasında hızlı dönüşüm yapar.\n"
            "8. Görsel Bağlantılarını Eşleştir:\n"
            "   Ayrı dosyadaki görsel bağlantılarını Excel/CSV içindeki anahtar sütuna ekler.\n"
            "9. QR Kod Üreticisi:\n"
            "   QR kodlarını PNG olarak veya isteğe bağlı alt metinli Dymo etiketleri şeklinde oluşturur.\n"
            "10. Barkod Üreticisi:\n"
            "   Popüler formatlarda barkodları görsel olarak veya yazdırmaya hazır Dymo etiketleri olarak üretir.\n"
            "--- ÖNERİLEN EKLEMELER ---\n"
            "• Yükleme öncesi başlık, etiket ve fiyatları normalize eden Shopify ürün verisi temizleyici.\n"
            "• Yerel CLI araçlarına veya web API'lerine bağlanan arka plan kaldırma yardımcısı.\n"
            "• Eksik medya, geçersiz UPC veya kopya handle'ları işaretleyen toplu CSV kalite kontrolcüsü.\n"
            "---------------------------------\n"
            "Hazırlayan: Hakan Akaslan\n"
        ),
        "warning_title": "Uyarı",
        "settings_empty": "Kaynak ve hedef klasörler boş olamaz.",
        "log_settings_saved": "✅ Ayarlar settings.json dosyasına kaydedildi",
        "success_title": "Başarılı",
        "settings_saved_popup": "Klasör ayarları kaydedildi.",
        "error_title": "Hata",
        "missing_required_fields": "Lütfen Kaynak, Hedef ve Numara Dosyasını belirtin.",
        "invalid_folder": "Lütfen geçerli bir klasör seçin.",
        "resize_value_error": "Boyutlandırma değerleri ve kalite geçerli sayı olmalıdır.",
        "select_file": "Lütfen bir dosya seçin.",
        "missing_file_and_column": "Lütfen bir dosya seçin ve bir sütun belirtin.",
        "missing_inputs": "Lütfen tüm dosya yollarını ve sütun adını doldurun.",
        "data_filename_required": "Veri ve dosya adı gereklidir.",
        "complete_title": "Tamamlandı",
    },
}

SUPPORTED_LANGUAGES = tuple(TEXTS.keys())


def sanitize_language(code: str | None) -> str:
    """Return a supported language code, defaulting to English."""
    if not code:
        return "en"
    return code if code in TEXTS else "en"


def display_for_language(code: str) -> str:
    """Return the human friendly label for a language code."""
    return LANGUAGE_DISPLAY.get(sanitize_language(code), LANGUAGE_DISPLAY["en"])


def translate(language_code: str, key: str, **kwargs) -> str:
    """Translate the given key for the current language."""
    lang = sanitize_language(language_code)
    template = TEXTS.get(lang, TEXTS["en"]).get(key, TEXTS["en"].get(key, key))
    if kwargs:
        return template.format(**kwargs)
    return template

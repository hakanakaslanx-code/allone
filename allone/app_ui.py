"""Main tkinter user interface for the desktop utility application."""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os
import math
from dataclasses import dataclass
from typing import Callable, List, Optional, Set, Tuple

import pandas as pd
import requests

import numpy as np

from PIL import Image, ImageTk

from print_service import SharedLabelPrinterServer, resolve_local_ip

from settings_manager import load_settings, save_settings
from updater import check_for_updates
import backend_logic as backend
from wayfair_formatter import WayfairFormatter

__version__ = "4.4.6"

translations = {
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
        "Compact Mode": "Compact Mode",
        "Zoom": "Zoom",
        "Advanced Settings": "Advanced Settings",
        "Show Sidebar": "Show Sidebar",
        "Hide Sidebar": "Hide Sidebar",
        "Sections": "Sections",
        "Automatic compact mode enabled for small screens.": "Automatic compact mode enabled for small screens.",
        "Language changed to {language}.": "Language changed to {language}.",
        "1. Copy/Move Files by List": "1. Copy/Move Files by List",
        "View in Room": "View in Room",
        "Source Folder:": "Source Folder:",
        "Target Folder:": "Target Folder:",
        "Numbers File (List):": "Numbers File (List):",
        "Browse...": "Browse...",
        "Browse": "Browse",
        "(Select)": "(Select)",
        "Image Files": "Image Files",
        "1. Choose Excel File": "1. Choose Excel File",
        "2. Columns": "2. Columns",
        "3. Map Wayfair Fields": "3. Map Wayfair Fields",
        "Width:": "Width:",
        "Length:": "Length:",
        "Load Mapping": "Load Mapping",
        "Save Mapping": "Save Mapping",
        "Create File": "Create File",
        "Select Excel File": "Select Excel File",
        "Excel Files": "Excel Files",
        "Excel file could not be read: {error}": "Excel file could not be read: {error}",
        "{count} columns loaded.": "{count} columns loaded.",
        "Please choose an Excel file first.": "Please choose an Excel file first.",
        "Missing Fields": "Missing Fields",
        "Please map the following fields:\n{fields}": "Please map the following fields:\n{fields}",
        "Size (Width/Length must be selected)": "Size (Width/Length must be selected)",
        "File could not be saved: {error}": "File could not be saved: {error}",
        "Wayfair formatted file saved to {path}.": "Wayfair formatted file saved to {path}.",
        "Missing Required Cells": "Missing Required Cells",
        "Some required cells are empty. See below for details.": "Some required cells are empty. See below for details.",
        "Wayfair formatted file is ready.": "Wayfair formatted file is ready.",
        "Missing fields: {fields}": "Missing fields: {fields}",
        "All required fields look filled.": "All required fields look filled.",
        "JSON": "JSON",
        "Mapping could not be saved: {error}": "Mapping could not be saved: {error}",
        "Mapping saved as {filename}.": "Mapping saved as {filename}.",
        "Mapping could not be loaded: {error}": "Mapping could not be loaded: {error}",
        "Mapping {filename} loaded.": "Mapping {filename} loaded.",
        "Width × Length": "Width × Length",
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
        "Room Image:": "Room Image:",
        "Rug Image:": "Rug Image:",
        "Resize Rug (%):": "Resize Rug (%):",
        "Yere Yayma (%):": "Yere Yayma (%):",
        "Rug Transparency:": "Rug Transparency:",
        "Lay Rug (90°)": "Lay Rug (90°)",
        "Generate Preview": "Generate Preview",
        "Save Image": "Save Image",
        "Manual Place Rug": "Halıyı Elle Yerleştir (4 Nokta)",
        "Preview will appear here.": "Preview will appear here.",
        "View in Room Controls": "Canvas controls: Left click and drag to move, mouse wheel to scale, right click and drag to rotate.",
        "Please select both room and rug images.": "Please select both room and rug images.",
        "Could not open selected images: {error}": "Could not open selected images: {error}",
        "Preview image saved to {path}.": "Preview image saved to {path}.",
        "No preview available. Please generate a preview first.": "No preview available. Please generate a preview first.",
        "Manual Prompt 1": "Üst sol köşeyi seçin",
        "Manual Prompt 2": "Üst sağ köşeyi seçin",
        "Manual Prompt 3": "Alt sağ köşeyi seçin",
        "Manual Prompt 4": "Alt sol köşeyi seçin",
        "Manual Placement Complete": "Halı yerleşti. Taşıyabilir veya ölçekleyebilirsiniz.",
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
        "Rug No Checker": "Rug No Checker",
        "Mode:": "Mode:",
        "Batch Comparison": "Batch Comparison",
        "Manual Search": "Manual Search",
        "Sold List File:": "Sold List File:",
        "Master List File:": "Master List File:",
        "Start Comparison": "Start Comparison",
        "Comparison Results:": "Comparison Results:",
        "Status": "Status",
        "Rug No": "Rug No",
        "Rug No Check": "Rug No Check",
        "Inventory List File:": "Inventory List File:",
        "Check Rug Nos": "Check Rug Nos",
        "Results:": "Results:",
        "RUG_NO_CONTROL_FOUND": "Found",
        "RUG_NO_CONTROL_NOT_FOUND": "Not Found",
        "FOUND": "FOUND",
        "MISSING": "MISSING",
        "Found: {found} | Missing: {missing}": "Found: {found} | Missing: {missing}",
        "Please select both Sold List and Master List files.": "Please select both Sold List and Master List files.",
        "Rug number comparison completed.": "Rug number comparison completed.",
        "Enter Rug No:": "Enter Rug No:",
        "Search": "Search",
        "Please select a Master List file.": "Please select a Master List file.",
        "Please enter a Rug No.": "Please enter a Rug No.",
        "Rug No {number} found in master list.": "Rug No {number} found in master list.",
        "Rug No {number} not found in master list.": "Rug No {number} not found in master list.",
        "Manual Search History:": "Manual Search History:",
        "No recent searches yet.": "No recent searches yet.",
        "Found": "Found",
        "Not Found": "Not Found",
        "Please select both Sold and Inventory files.": "Please select both Sold and Inventory files.",
        "Could not find a Rug No column in the selected file.": "Could not find a Rug No column in the selected file.",
        "Could not read the selected file: {error}": "Could not read the selected file: {error}",
        "Rug No control completed.": "Rug No control completed.",
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
        "Rinven Tag": "Rinven Tag",
        "Collection Name:": "Collection Name:",
        "Design:": "Design:",
        "Color:": "Color:",
        "Size:": "Size:",
        "Origin:": "Origin:",
        "Style:": "Style:",
        "Content:": "Content:",
        "Type:": "Type:",
        "Rug #:": "Rug #:",
        "Include Barcode": "Include Barcode",
        "Barcode Data:": "Barcode Data:",
        "Generate Rinven Tag": "Generate Rinven Tag",
        "Shared Label Printer": "Shared Label Printer",
        "SHARED_PRINTER_DESCRIPTION": (
            "Expose the locally connected DYMO LabelWriter 450 to other computers on your Wi-Fi/LAN.\n"
            "Start sharing to run the embedded Flask server and accept POST /print jobs with the token below."
        ),
        "Authorization Token:": "Authorization Token:",
        "Listen Port:": "Listen Port:",
        "Server Status: ⚪ Stopped": "Server Status: ⚪ Stopped",
        "Server Status: 🟢 Running": "Server Status: 🟢 Running",
        "Server Status: ⏳ Checking...": "Server Status: ⏳ Checking...",
        "Start Sharing": "Start Sharing",
        "Stop Sharing": "Stop Sharing",
        "Check Status": "Check Status",
        "SHARED_PRINTER_STARTED": "Shared label printer server started on {host}:{port}.",
        "SHARED_PRINTER_STOPPED": "Shared label printer server stopped.",
        "SHARED_PRINTER_START_FAILED": "Failed to start sharing: {error}",
        "SHARED_PRINTER_STATUS_FAILED": "Status request failed: {error}",
        "SHARED_PRINTER_TOKEN_REQUIRED": "Please enter an authorization token.",
        "SHARED_PRINTER_AUTOSTARTED": "Shared label printer server auto-started on {host}:{port}.",
        "SHARED_PRINTER_AUTOSTART_FAILED": "Automatic start failed: {error}",
        "SHARED_PRINTER_STATUS_DETAIL": "Server Status: 🟢 Running — {host}:{port}",
        "SHARED_PRINTER_HELP_TEXT": (
            "Other PCs on this same Wi-Fi / LAN can print to this label printer by sending a POST /print request to http://{host}:{port}/print with the same bearer token. Do not expose this port to the internet."
        ),
        "SHARED_PRINTER_DISABLED": "Shared label printer sharing is currently disabled.",
        "SHARED_PRINTER_NOT_READY": "Shared label printer server is not ready yet.",
        "Server port is not configured.": "Server port is not configured.",
        "Please enter a valid port number.": "Please enter a valid port number.",
        "Please fill in all Rinven Tag fields.": "Please fill in all Rinven Tag fields.",
        "Barcode data is required when barcode is enabled.": "Barcode data is required when barcode is enabled.",
        "Filename is required.": "Filename is required.",
        "win32print is only available on Windows.": "win32print is only available on Windows.",
        "win32print module could not be loaded. Please check the pywin32 installation.": "win32print module could not be loaded. Please check the pywin32 installation.",
        "Server token is not configured.": "Server token is not configured.",
        "Invalid or missing authorization token.": "Invalid or missing authorization token.",
        "No file found in request.": "No file found in request.",
        "No valid filename provided.": "No valid filename provided.",
        "Printer name could not be determined.": "Printer name could not be determined.",
        "File content is empty; cannot print.": "File content is empty; cannot print.",
        "Shared printer server error: {error}": "Shared printer server error: {error}",
        "Server is already running.": "Server is already running.",
        "Authorization token cannot be empty.": "Authorization token cannot be empty.",
        "Port {port} is not available: {error}": "Port {port} is not available: {error}",
        "Print error: %s": "Print error: %s",
        "Check for Updates": "Check for Updates",
        "Warning": "Warning",
        "Information": "Information",
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
        "Network Printers": "Network Printers",
        "NETWORK_PRINTERS_DESCRIPTION": (
            "Discover printers shared over your LAN. Select a local or remote printer "
            "and send files through the AllOne Tools print service."
        ),
        "Select Printer:": "Select Printer:",
        "Refresh Local Printers": "Refresh Local Printers",
        "Printer File:": "Printer File:",
        "Browse File": "Browse File",
        "Send Print Job": "Send Print Job",
        "Discovered Network Printers": "Discovered Network Printers",
        "Printer Name": "Printer Name",
        "Host Computer": "Host Computer",
        "IP Address": "IP Address",
        "Port": "Port",
        "Origin": "Origin",
        "Local": "Local",
        "Remote": "Remote",
        "No printers available.": "No printers available.",
        "Please select a printer.": "Please select a printer.",
        "Please select a file to print.": "Please select a file to print.",
        "PRINT_JOB_SENT": "Print job sent successfully.",
        "PRINT_JOB_FAILED": "Failed to send print job: {error}",
        "DISCOVERY_UNAVAILABLE": "Printer discovery service unavailable.",
        "Printer backend unavailable.": "Printer backend unavailable.",
        "Network printer discovery is not available on this platform.": "Network printer discovery is not available on this platform.",
        "Select a file to send to the printer.": "Select a file to send to the printer.",
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
        "Compact Mode": "Kompakt Mod",
        "Zoom": "Yakınlaştırma",
        "Advanced Settings": "Gelişmiş Ayarlar",
        "Show Sidebar": "Yan Paneli Göster",
        "Hide Sidebar": "Yan Paneli Gizle",
        "Sections": "Bölümler",
        "Automatic compact mode enabled for small screens.": "Küçük ekranlar için otomatik kompakt mod etkinleştirildi.",
        "Language changed to {language}.": "Dil {language} olarak değiştirildi.",
        "1. Copy/Move Files by List": "1. Listeye Göre Dosya Kopyala/Taşı",
        "View in Room": "Odanızda Görüntüle",
        "Source Folder:": "Kaynak Klasör:",
        "Target Folder:": "Hedef Klasör:",
        "Numbers File (List):": "Numara Dosyası (Liste):",
        "Browse...": "Gözat...",
        "Browse": "Gözat",
        "(Select)": "(Seçiniz)",
        "Image Files": "Görsel Dosyaları",
        "1. Choose Excel File": "1. Excel Dosyası Seç",
        "2. Columns": "2. Sütunlar",
        "3. Map Wayfair Fields": "3. Wayfair Alan Eşlemesi",
        "Width:": "Genişlik:",
        "Length:": "Uzunluk:",
        "Load Mapping": "Mapping Yükle",
        "Save Mapping": "Mapping Kaydet",
        "Create File": "Dosya Oluştur",
        "Select Excel File": "Excel Dosyası Seç",
        "Excel Files": "Excel Dosyaları",
        "Excel file could not be read: {error}": "Excel dosyası okunamadı: {error}",
        "{count} columns loaded.": "{count} sütun yüklendi.",
        "Please choose an Excel file first.": "Lütfen önce bir Excel dosyası seçin.",
        "Missing Fields": "Eksik Alanlar",
        "Please map the following fields:\n{fields}": "Lütfen aşağıdaki alanlar için eşleme yapın:\n{fields}",
        "Size (Width/Length must be selected)": "Size (Width/Length seçilmeli)",
        "File could not be saved: {error}": "Dosya kaydedilemedi: {error}",
        "Wayfair formatted file saved to {path}.": "Wayfair formatlı dosya {path} konumuna kaydedildi.",
        "Missing Required Cells": "Eksik Zorunlu Hücreler",
        "Some required cells are empty. See below for details.": "Bazı zorunlu hücreler boş. Ayrıntılar için aşağıya bakın.",
        "Wayfair formatted file is ready.": "Wayfair formatında dosya hazır.",
        "Missing fields: {fields}": "Eksik alanlar: {fields}",
        "All required fields look filled.": "Tüm zorunlu alanlar dolu görünüyor.",
        "JSON": "JSON",
        "Mapping could not be saved: {error}": "Mapping kaydedilemedi: {error}",
        "Mapping saved as {filename}.": "Mapping {filename} olarak kaydedildi.",
        "Mapping could not be loaded: {error}": "Mapping yüklenemedi: {error}",
        "Mapping {filename} loaded.": "Mapping {filename} yüklendi.",
        "Width × Length": "Genişlik × Uzunluk",
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
        "Room Image:": "Oda Görseli:",
        "Rug Image:": "Halı Görseli:",
        "Resize Rug (%):": "Halı Boyutu (%):",
        "Yere Yayma (%):": "Yere Yayma (%):",
        "Rug Transparency:": "Halı Saydamlığı:",
        "Lay Rug (90°)": "Halıyı Yatır (90°)",
        "Generate Preview": "Önizleme Oluştur",
        "Save Image": "Görseli Kaydet",
        "Manual Place Rug": "Halıyı Elle Yerleştir (4 Nokta)",
        "Preview will appear here.": "Önizleme burada görünecek.",
        "View in Room Controls": "Kontroller: Taşımak için sol tıkla sürükleyin, ölçek için tekerleği çevirin, döndürmek için sağ tıkla sürükleyin.",
        "Please select both room and rug images.": "Lütfen hem oda hem halı görsellerini seçin.",
        "Could not open selected images: {error}": "Seçilen görseller açılamadı: {error}",
        "Preview image saved to {path}.": "Önizleme görseli {path} konumuna kaydedildi.",
        "No preview available. Please generate a preview first.": "Önizleme yok. Lütfen önce bir önizleme oluşturun.",
        "Manual Prompt 1": "Üst sol köşeyi seçin",
        "Manual Prompt 2": "Üst sağ köşeyi seçin",
        "Manual Prompt 3": "Alt sağ köşeyi seçin",
        "Manual Prompt 4": "Alt sol köşeyi seçin",
        "Manual Placement Complete": "Halı yerleşti. Taşıyabilir veya ölçekleyebilirsiniz.",
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
        "Rug No Checker": "Rug No Kontrolü",
        "Rug No Check": "Rug No Kontrol",
        "Mode:": "Mod:",
        "Batch Comparison": "Toplu Karşılaştırma",
        "Manual Search": "Manuel Arama",
        "Sold List File:": "Sold List Dosyası:",
        "Inventory List File:": "Envanter Listesi Dosyası:",
        "Master List File:": "Master List Dosyası:",
        "Start Comparison": "Karşılaştırmayı Başlat",
        "Comparison Results:": "Karşılaştırma Sonuçları:",
        "Status": "Durum",
        "Rug No": "Rug No",
        "Check Rug Nos": "Rug No Kontrol Et",
        "Results:": "Sonuçlar:",
        "RUG_NO_CONTROL_FOUND": "Bulundu",
        "RUG_NO_CONTROL_NOT_FOUND": "Yok",
        "FOUND": "BULUNDU",
        "MISSING": "EKSİK",
        "Found: {found} | Missing: {missing}": "Bulundu: {found} | Eksik: {missing}",
        "Please select both Sold List and Master List files.": "Lütfen hem Sold List hem de Master List dosyalarını seçin.",
        "Rug number comparison completed.": "Rug No karşılaştırması tamamlandı.",
        "Enter Rug No:": "Rug No Girin:",
        "Search": "Ara",
        "Please select a Master List file.": "Lütfen bir Master List dosyası seçin.",
        "Please enter a Rug No.": "Lütfen bir Rug No girin.",
        "Rug No {number} found in master list.": "Rug No {number} master listede bulundu.",
        "Rug No {number} not found in master list.": "Rug No {number} master listede bulunamadı.",
        "Manual Search History:": "Manuel Arama Geçmişi:",
        "No recent searches yet.": "Henüz arama geçmişi yok.",
        "Found": "Bulundu",
        "Not Found": "Bulunamadı",
        "Please select both Sold and Inventory files.": "Lütfen Satılanlar ve Envanter dosyalarını seçin.",
        "Could not find a Rug No column in the selected file.": "Seçilen dosyada bir Halı No sütunu bulunamadı.",
        "Could not read the selected file: {error}": "Seçilen dosya okunamadı: {error}",
        "Rug No control completed.": "Rug No kontrolü tamamlandı.",
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
        "Rinven Tag": "Rinven Etiketi",
        "Collection Name:": "Koleksiyon Adı:",
        "Design:": "Desen:",
        "Color:": "Renk:",
        "Size:": "Boyut:",
        "Origin:": "Menşei:",
        "Style:": "Stil:",
        "Content:": "İçerik:",
        "Type:": "Tür:",
        "Rug #:": "Halı No:",
        "Include Barcode": "Barkodu Dahil Et",
        "Barcode Data:": "Barkod Verisi:",
        "Generate Rinven Tag": "Rinven Etiketi Oluştur",
        "Shared Label Printer": "Paylaşılan Etiket Yazıcısı",
        "SHARED_PRINTER_DESCRIPTION": (
            "Yerel olarak bağlı DYMO LabelWriter 450 yazıcısını Wi-Fi/LAN üzerindeki diğer bilgisayarlarla paylaşın.\n"
            "Aşağıdaki jetonla gömülü Flask sunucusunu başlatın ve POST /print isteklerini kabul edin."
        ),
        "Authorization Token:": "Yetkilendirme Jetonu:",
        "Listen Port:": "Dinleme Portu:",
        "Server Status: ⚪ Stopped": "Sunucu Durumu: ⚪ Kapalı",
        "Server Status: 🟢 Running": "Sunucu Durumu: 🟢 Çalışıyor",
        "Server Status: ⏳ Checking...": "Sunucu Durumu: ⏳ Kontrol ediliyor...",
        "Start Sharing": "Paylaşımı Başlat",
        "Stop Sharing": "Paylaşımı Durdur",
        "Check Status": "Durumu Kontrol Et",
        "SHARED_PRINTER_STARTED": "Etiket yazıcısı paylaşımı {host}:{port} adresinde başlatıldı.",
        "SHARED_PRINTER_STOPPED": "Etiket yazıcısı paylaşımı durduruldu.",
        "SHARED_PRINTER_START_FAILED": "Paylaşım başlatılamadı: {error}",
        "SHARED_PRINTER_STATUS_FAILED": "Durum isteği başarısız: {error}",
        "SHARED_PRINTER_TOKEN_REQUIRED": "Lütfen bir yetkilendirme jetonu girin.",
        "SHARED_PRINTER_AUTOSTARTED": "Etiket yazıcısı paylaşımı otomatik olarak {host}:{port} adresinde başlatıldı.",
        "SHARED_PRINTER_AUTOSTART_FAILED": "Otomatik başlatma başarısız oldu: {error}",
        "SHARED_PRINTER_STATUS_DETAIL": "Sunucu Durumu: 🟢 Çalışıyor — {host}:{port}",
        "SHARED_PRINTER_HELP_TEXT": (
            "Aynı Wi-Fi / LAN içindeki diğer bilgisayarlar http://{host}:{port}/print adresine aynı bearer jetonuyla POST /print isteği göndererek bu yazıcıya çıktı alabilir. Bu portu internete açmayın."
        ),
        "SHARED_PRINTER_DISABLED": "Etiket yazıcısı paylaşımı şu anda devre dışı.",
        "SHARED_PRINTER_NOT_READY": "Paylaşılan etiket yazıcısı sunucusu henüz hazır değil.",
        "Server port is not configured.": "Sunucu portu yapılandırılmadı.",
        "Please enter a valid port number.": "Lütfen geçerli bir port numarası girin.",
        "Please fill in all Rinven Tag fields.": "Lütfen tüm Rinven Etiketi alanlarını doldurun.",
        "Barcode data is required when barcode is enabled.": "Barkod etkinleştirildiğinde barkod verisi gereklidir.",
        "Filename is required.": "Dosya adı gereklidir.",
        "win32print is only available on Windows.": "win32print sadece Windows üzerinde kullanılabilir.",
        "win32print module could not be loaded. Please check the pywin32 installation.": "win32print modülü yüklenemedi. Lütfen pywin32 kurulumunu kontrol edin.",
        "Server token is not configured.": "Sunucu jetonu yapılandırılmamış.",
        "Invalid or missing authorization token.": "Geçersiz veya eksik yetkilendirme jetonu.",
        "No file found in request.": "Yüklenecek dosya bulunamadı.",
        "No valid filename provided.": "Geçerli bir dosya adı gönderilmedi.",
        "Printer name could not be determined.": "Yazıcı adı bulunamadı.",
        "File content is empty; cannot print.": "Dosya içeriği boş olduğu için yazdırma yapılamadı.",
        "Shared printer server error: {error}": "Paylaşılan yazıcı sunucusu hata verdi: {error}",
        "Server is already running.": "Sunucu zaten çalışıyor.",
        "Authorization token cannot be empty.": "Yetkilendirme jetonu boş olamaz.",
        "Port {port} is not available: {error}": "Port {port} kullanılamıyor: {error}",
        "Print error: %s": "Yazdırma sırasında hata oluştu: %s",
        "Check for Updates": "Güncellemeleri Kontrol Et",
        "Warning": "Uyarı",
        "Information": "Bilgi",
        "Source and Target folders cannot be empty.": "Kaynak ve hedef klasörler boş olamaz.",
        "✅ Settings saved to settings.json": "✅ Ayarlar settings.json dosyasına kaydedildi",
        "Success": "Başarılı",
        "Folder settings have been saved.": "Klasör ayarları kaydedildi.",
        "Error": "Hata",
        "Please specify Source, Target, and Numbers File.": "Lütfen Kaynak, Hedef ve Numara Dosyasını belirtin.",
        "Network Printers": "Ağ Yazıcıları",
        "NETWORK_PRINTERS_DESCRIPTION": (
            "Yerel ağınızda paylaşılan yazıcıları keşfedin. Yerel veya uzak bir yazıcı "
            "seçip AllOne Tools yazdırma servisi üzerinden dosya gönderin."
        ),
        "Select Printer:": "Yazıcı Seç:",
        "Refresh Local Printers": "Yerel Yazıcıları Yenile",
        "Printer File:": "Yazdırılacak Dosya:",
        "Browse File": "Dosya Seç",
        "Send Print Job": "Yazdırmayı Gönder",
        "Discovered Network Printers": "Keşfedilen Ağ Yazıcıları",
        "Printer Name": "Yazıcı Adı",
        "Host Computer": "Bilgisayar",
        "IP Address": "IP Adresi",
        "Port": "Port",
        "Origin": "Kaynak",
        "Local": "Yerel",
        "Remote": "Uzak",
        "No printers available.": "Kullanılabilir yazıcı yok.",
        "Please select a printer.": "Lütfen bir yazıcı seçin.",
        "Please select a file to print.": "Lütfen yazdırılacak bir dosya seçin.",
        "PRINT_JOB_SENT": "Yazdırma isteği gönderildi.",
        "PRINT_JOB_FAILED": "Yazdırma isteği başarısız: {error}",
        "DISCOVERY_UNAVAILABLE": "Yazıcı keşif servisi kullanılamıyor.",
        "Printer backend unavailable.": "Yazıcı altyapısı kullanılamıyor.",
        "Network printer discovery is not available on this platform.": "Bu platformda ağ yazıcı keşfi desteklenmiyor.",
        "Select a file to send to the printer.": "Yazıcıya gönderilecek dosyayı seçin.",
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


RUG_NO_CONTROL_COLUMNS = ["Rug No", "RugNo", "RugNo#", "SKU", "Sku"]


class ScrollableTab(ttk.Frame):
    """Wraps a frame within a canvas to provide per-tab scrolling."""

    def __init__(self, parent: ttk.Notebook):
        super().__init__(parent)

        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.interior = ttk.Frame(self.canvas, style="TFrame", padding=4)
        self._window_id = self.canvas.create_window((0, 0), window=self.interior, anchor="nw")

        self.interior.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.interior.bind("<Enter>", self._bind_mousewheel)
        self.interior.bind("<Leave>", self._unbind_mousewheel)

    def _on_frame_configure(self, event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self._window_id, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> str:
        if getattr(event, "delta", 0):
            direction = -1 if event.delta > 0 else 1
        else:
            direction = -1 if getattr(event, "num", 0) == 4 else 1
        self.canvas.yview_scroll(direction, "units")
        return "break"

    def _bind_mousewheel(self, _event: tk.Event) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event: tk.Event) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

PANEL_INFO = {
    "en": {
        "1. Copy/Move Files by List": "Quickly copy or move files referenced by a spreadsheet of item numbers.",
        "2. Convert HEIC to JPG": "Convert entire folders of HEIC photos into widely compatible JPG images.",
        "3. Batch Image Resizer": "Resize and compress images in bulk using width- or percentage-based rules.",
        "View in Room": "Preview rugs inside a selected room photo with scaling and transparency controls.",
        "4. Format Numbers from File": "Clean and format numbers from Excel, CSV or text files for exports.",
        "5. Rug Size Calculator (Single)": "Calculate exact square footage and square meters for a single rug size.",
        "6. BULK Process Rug Sizes from File": "Normalize every rug size inside a spreadsheet using your chosen column.",
        "Rug No Checker": "Compare sold and master lists or search manually to verify rug numbers.",
        "Rug No Check": "Check sold rug numbers against an inventory file and report their availability.",
        "7. Unit Converter": "Convert between popular measurements such as centimeters, inches and feet.",
        "8. Match Image Links": "Attach hosted image URLs to product rows by matching a shared key column.",
        "8. QR Code Generator": "Generate QR codes for web links or label printers in just a few clicks.",
        "9. Barcode Generator": "Create printable barcodes in multiple formats, including DYMO labels.",
        "Wayfair Export Formatter": "Prepare a Wayfair-ready Excel file by mapping columns and validating required fields.",
        "Rinven Tag": "Design branded Rinven tags with collection details and optional barcode.",
        "Shared Label Printer": "Share your local DYMO printer securely with other devices on the network.",
        "Help & About": "Review update status, helpful links and support information for the app.",
    },
    "tr": {
        "1. Copy/Move Files by List": "Numara listesindeki kayıtlara göre dosyaları hızlıca kopyalayın veya taşıyın.",
        "2. Convert HEIC to JPG": "Tüm HEIC fotoğraflarını tek seferde yaygın kullanılan JPG formatına dönüştürür.",
        "3. Batch Image Resizer": "Görselleri genişliğe ya da yüzdeye göre toplu biçimde yeniden boyutlandırıp sıkıştırır.",
        "View in Room": "Seçtiğiniz oda fotoğrafında halıyı ölçek ve saydamlıkla yerleştirerek önizleyin.",
        "4. Format Numbers from File": "Excel, CSV veya TXT dosyalarındaki sayıları dışa aktarıma uygun biçimde temizler.",
        "5. Rug Size Calculator (Single)": "Tek bir halı ölçüsünün metrekare ve fit değerlerini anında hesaplar.",
        "6. BULK Process Rug Sizes from File": "Seçtiğiniz sütundaki tüm halı ölçülerini standart forma dönüştürür.",
        "Rug No Checker": "Satış ve ana listeleri karşılaştırarak halı numaralarını doğrular veya manuel arama yapar.",
        "Rug No Check": "Satılan halı numaralarını envanter dosyasıyla hızlıca karşılaştırıp durumlarını raporlar.",
        "7. Unit Converter": "Sık kullanılan ölçü birimlerini (cm, inç, feet vb.) hızlıca çevirir.",
        "8. Match Image Links": "Paylaşılan anahtar sütunu kullanarak ürün satırlarına görsel bağlantıları ekler.",
        "8. QR Code Generator": "Web bağlantıları veya etiket yazıcıları için birkaç tıklamayla QR kodu üretir.",
        "9. Barcode Generator": "PNG veya DYMO dahil birden çok formatta baskıya hazır barkod oluşturur.",
        "Wayfair Export Formatter": "Wayfair'e uygun Excel çıktısını sütun eşleştirme ve doğrulama ile hazırlayın.",
        "Rinven Tag": "Koleksiyon bilgileri ve isteğe bağlı barkod içeren Rinven etiketleri tasarlar.",
        "Shared Label Printer": "Yerel DYMO yazıcınızı ağdaki diğer cihazlarla güvenle paylaşmanızı sağlar.",
        "Help & About": "Uygulama sürümünü, rehberleri ve destek bağlantılarını tek yerde gösterir.",
    },
}

DYMO_LABELS = {
    'Address (30252)': {'w_in': 3.5, 'h_in': 1.125},
    'Shipping (30256)': {'w_in': 4.0, 'h_in': 2.3125},
    'Small Multipurpose (30336)': {'w_in': 2.125, 'h_in': 1.0},
    'File Folder (30258)': {'w_in': 3.5, 'h_in': 0.5625},
}


@dataclass
class RugWarpResult:
    image: Image.Image
    offset: Tuple[float, float]
    size: Tuple[int, int]
    polygon: List[Tuple[float, float]]


class ToolApp(tk.Tk):
    """Main application window that builds the entire tkinter interface."""

    def __init__(self):
        super().__init__()

        self.settings = load_settings()
        self.settings.setdefault("rinven_history", {})
        legacy_print = self.settings.get("print_server", {})
        shared_settings = self.settings.setdefault("shared_label_printer", {})
        shared_settings.setdefault("token", legacy_print.get("token", "change-me"))
        shared_settings.setdefault("port", legacy_print.get("port", 5151))
        shared_settings.setdefault("autostart_on_launch", False)
        if "print_server" in self.settings:
            # Eski ayar anahtarını temizleyerek tek bir kaynaktan devam ediyoruz.
            self.settings.pop("print_server", None)
            save_settings(self.settings)
        self.language = self.settings.get("language", "en")
        if self.language not in translations:
            self.language = "en"

        self.translatable_widgets = []
        self.section_descriptions = []
        self.notebook_tabs = []
        self.sidebar_nav = []
        self.advanced_cards = []
        self.view_in_room_preview_photo: Optional[ImageTk.PhotoImage] = None
        self.view_in_room_preview_image: Optional[Image.Image] = None
        self.view_in_room_preview_has_image = False
        self.view_in_room_room_image: Optional[Image.Image] = None
        self.view_in_room_rug_original: Optional[Image.Image] = None
        self.view_in_room_rug_scale: float = 1.0
        self.view_in_room_rug_angle: float = 0.0
        self.view_in_room_rug_center: Optional[Tuple[float, float]] = None
        self.view_in_room_display_scale: float = 1.0
        self.view_in_room_canvas_room_photo: Optional[ImageTk.PhotoImage] = None
        self.view_in_room_canvas_rug_photo: Optional[ImageTk.PhotoImage] = None
        self.view_in_room_canvas_room_item = None
        self.view_in_room_canvas_rug_item = None
        self.view_in_room_canvas_message = None
        self.view_in_room_drag_mode: Optional[str] = None
        self.view_in_room_drag_offset: Tuple[float, float] = (0.0, 0.0)
        self.view_in_room_rotation_reference: float = 0.0
        self.view_in_room_rotation_start_angle: float = 0.0
        self.view_in_room_scale_reference_distance: float = 0.0
        self.view_in_room_scale_start_value: float = 1.0
        self.view_in_room_rug_display_bbox: Optional[Tuple[float, float, float, float]] = None
        self.view_in_room_rug_display_center: Optional[Tuple[float, float]] = None
        self.view_in_room_display_size: Tuple[int, int] = (720, 540)
        self.view_in_room_canvas_size: Tuple[int, int] = (720, 540)
        self.view_in_room_perspective_top_scale: float = 0.6
        self.view_in_room_manual_active: bool = False
        self.view_in_room_manual_points: List[Tuple[float, float]] = []
        self.view_in_room_manual_relative_polygon: Optional[List[Tuple[float, float]]] = None
        self.view_in_room_manual_prompt_var: Optional[tk.StringVar] = None
        self.view_in_room_manual_button: Optional[ttk.Button] = None
        self.view_in_room_manual_prompt_label: Optional[ttk.Label] = None

        self.language_options = {"en": "English", "tr": "Turkish"}
        self.language_var = tk.StringVar(
            value=self.tr(self.language_options.get(self.language, "English"))
        )
        self._updating_language_selector = False

        self.ui_preferences = self.settings.setdefault("ui_preferences", {})
        self._base_named_font_sizes = {}
        for name in (
            "TkDefaultFont",
            "TkTextFont",
            "TkMenuFont",
            "TkHeadingFont",
            "TkCaptionFont",
            "TkSmallCaptionFont",
            "TkFixedFont",
            "TkIconFont",
            "TkTooltipFont",
        ):
            try:
                font_obj = tkfont.nametofont(name)
                self._base_named_font_sizes.setdefault(name, font_obj.cget("size"))
            except tk.TclError:
                continue

        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.small_screen = screen_width < 1366 or screen_height < 900

        base_width, base_height = 900, 750
        margin = 80  # Leave a small margin so the window isn't flush with edges.
        min_width, min_height = 600, 500
        available_width = min(screen_width, max(min_width, screen_width - margin))
        available_height = min(screen_height, max(min_height, screen_height - margin))
        target_width = min(max(base_width, int(screen_width * 0.9)), available_width)
        target_height = min(max(base_height, int(screen_height * 0.9)), available_height)

        stored_geometry = self.ui_preferences.get("window_geometry")
        if stored_geometry:
            try:
                self.geometry(stored_geometry)
            except tk.TclError:
                self.geometry(f"{target_width}x{target_height}")
        else:
            self.geometry(f"{target_width}x{target_height}")
        self.minsize(min(target_width, base_width), min(target_height, base_height))

        stored_compact = self.ui_preferences.get("compact_mode")
        if stored_compact is None:
            self.compact_mode = self.small_screen
            self._auto_compact_message = self.compact_mode
        else:
            self.compact_mode = bool(stored_compact)
            self._auto_compact_message = False

        self._zoom_levels = ["80%", "100%", "120%"]
        stored_zoom = self.ui_preferences.get("zoom_level", "100%")
        if stored_zoom not in self._zoom_levels:
            stored_zoom = "100%"
        self.zoom_level = stored_zoom
        self._zoom_factor = self._parse_zoom_level(self.zoom_level)
        self.sidebar_collapsed = bool(self.ui_preferences.get("sidebar_collapsed", False))
        self.show_advanced = bool(self.ui_preferences.get("show_advanced", True))

        self._update_named_fonts()
        self.setup_styles()
        self.create_header()
        self.create_view_toolbar()
        self.create_language_selector()

        self.shared_token_var = tk.StringVar(value=str(shared_settings.get("token", "change-me")))
        self.shared_port_var = tk.StringVar(value=str(shared_settings.get("port", 5151)))
        self.shared_status_var = tk.StringVar(value=self.tr("Server Status: ⚪ Stopped"))
        self.shared_status_state = "stopped"
        self.shared_status_host: Optional[str] = None
        self.shared_status_port: Optional[int] = None
        self.shared_printer_server = SharedLabelPrinterServer(self.log, translator=self.tr)
        self.shared_status_lock = threading.RLock()
        self.shared_port_var.trace_add("write", lambda *args: self._update_shared_help_text())

        self.main_body = ttk.Frame(self, style="TFrame")
        self.main_body.pack(pady=(5, 12), padx=12, fill="both", expand=True)
        self.main_body.columnconfigure(1, weight=1)
        self.main_body.rowconfigure(0, weight=1)

        self.sidebar_frame = ttk.Frame(self.main_body, style="Sidebar.TFrame", width=220)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        self.sidebar_frame.grid_propagate(False)

        sidebar_header = ttk.Frame(self.sidebar_frame, style="SidebarHeader.TFrame")
        sidebar_header.pack(fill="x", padx=8, pady=(8, 4))
        sidebar_label = ttk.Label(sidebar_header, text=self.tr("Sections"), style="Secondary.TLabel")
        sidebar_label.pack(anchor="w")
        self.register_widget(sidebar_label, "Sections")
        self.sidebar_header_label = sidebar_label

        self.sidebar_button_area = ttk.Frame(self.sidebar_frame, style="Sidebar.TFrame")
        self.sidebar_button_area.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.content_frame = ttk.Frame(self.main_body, style="TFrame")
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(0, weight=5)
        self.content_frame.rowconfigure(1, weight=2)

        self.section_notebook = ttk.Notebook(self.content_frame, style="TNotebook")
        self.section_notebook.grid(row=0, column=0, sticky="nsew")

        self.section_frames = {}
        for title in (
            "File & Image Tools",
            "View in Room",
            "Data & Calculation",
            "Rug No Check",
            "Code Generators",
            "Rinven Tag",
            "Shared Label Printer",
            "Help & About",
        ):
            tab = ScrollableTab(self.section_notebook)
            tab.interior.columnconfigure(0, weight=1)
            tab.interior.columnconfigure(1, weight=1)
            self.section_notebook.add(tab, text=self.tr(title))
            self.section_frames[title] = tab.interior
            self.notebook_tabs.append((tab, title))
            self._create_sidebar_button(title, tab)

        self.section_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self.rug_manual_history_entries = []
        self.rug_manual_last_result = None
        self.rug_comparison_results = None
        self.rug_control_results: List[Tuple[str, bool]] = []

        self.create_file_image_panels(self.section_frames["File & Image Tools"])
        self.create_view_in_room_tab(self.section_frames["View in Room"])
        self.create_data_calc_panels(self.section_frames["Data & Calculation"])
        self.create_rug_no_control_tab(self.section_frames["Rug No Check"])
        self.create_code_gen_panels(self.section_frames["Code Generators"])
        self.create_rinven_tag_panel(self.section_frames["Rinven Tag"])
        self.create_shared_printer_panel(self.section_frames["Shared Label Printer"])
        self.create_about_panel(self.section_frames["Help & About"])

        self.log_area = ScrolledText(self.content_frame, height=8)
        self.log_area.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self._apply_log_theme()

        self._apply_sidebar_visibility(initial=True)
        self._apply_advanced_visibility()

        self.refresh_translations()
        self.log(self.tr("Welcome to the Combined Utility Tool!"))
        if self._auto_compact_message:
            self.log(self.tr("Automatic compact mode enabled for small screens."))
            self._auto_compact_message = False

        self.run_in_thread(check_for_updates, self, self.log, __version__, silent=True)

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(0, self._maybe_auto_start_shared_printer)

    def _parse_zoom_level(self, value: str) -> float:
        mapping = {"80%": 0.8, "100%": 1.0, "120%": 1.2}
        return mapping.get(value, 1.0)

    def _density_multiplier(self) -> float:
        return 0.85 if self.compact_mode else 1.0

    def _pad_value(self, base: int, minimum: int = 2) -> int:
        return max(minimum, int(round(base * self._density_multiplier())))

    def _scaled_size(self, base: int) -> int:
        size = abs(base)
        size = max(6, int(round(size * self._zoom_factor)))
        if self.compact_mode:
            size = max(6, int(round(size * 0.9)))
        return size if base >= 0 else -size

    def _font(self, family: str, base: int, weight: Optional[str] = None) -> tuple:
        size = self._scaled_size(base)
        if weight:
            return (family, size, weight)
        return (family, size)

    def _update_named_fonts(self) -> None:
        try:
            self.tk.call("tk", "scaling", self._zoom_factor)
        except tk.TclError:
            pass
        for name, base_size in self._base_named_font_sizes.items():
            try:
                font_obj = tkfont.nametofont(name)
            except tk.TclError:
                continue
            scaled = self._scaled_size(base_size)
            if base_size < 0:
                font_obj.configure(size=-scaled)
            else:
                font_obj.configure(size=scaled)

    def _apply_log_theme(self) -> None:
        if not hasattr(self, "log_area"):
            return
        self.log_area.config(
            state=tk.DISABLED,
            background="#0b1120",
            foreground="#f1f5f9",
            insertbackground="#f1f5f9",
            font=("Cascadia Code", self._scaled_size(10)),
            relief="flat",
            borderwidth=0,
        )

    def _create_sidebar_button(self, title: str, tab: ttk.Frame) -> None:
        button = ttk.Button(
            self.sidebar_button_area,
            text=self.tr(title),
            style="Sidebar.TButton",
            command=lambda t=tab: self.section_notebook.select(t),
        )
        button.pack(fill="x", pady=(0, 4))
        self.register_widget(button, title)
        self.sidebar_nav.append((button, tab))

    def _update_nav_highlight(self) -> None:
        if not hasattr(self, "section_notebook"):
            return
        current = self.section_notebook.select()
        for button, tab in self.sidebar_nav:
            if str(tab) == current:
                button.configure(style="SidebarSelected.TButton")
            else:
                button.configure(style="Sidebar.TButton")

    def _on_tab_changed(self, _event=None) -> None:
        self._update_nav_highlight()

    def _apply_sidebar_visibility(self, initial: bool = False) -> None:
        if self.sidebar_collapsed:
            self.sidebar_frame.grid_remove()
        else:
            if initial:
                self.sidebar_frame.grid()  # ensure geometry manager remembers placement
            else:
                self.sidebar_frame.grid()
        self._update_sidebar_toggle_text()
        self._update_nav_highlight()

    def _update_sidebar_toggle_text(self) -> None:
        if hasattr(self, "sidebar_toggle"):
            key = "Show Sidebar" if self.sidebar_collapsed else "Hide Sidebar"
            self.sidebar_toggle.configure(text=self.tr(key))

    def _toggle_sidebar(self) -> None:
        self.sidebar_collapsed = not self.sidebar_collapsed
        self._apply_sidebar_visibility()
        self._save_ui_preferences()

    def _on_toggle_compact(self) -> None:
        self.compact_mode = bool(self.compact_var.get())
        self._update_named_fonts()
        self.setup_styles()
        self._apply_advanced_visibility()
        self._save_ui_preferences()

    def _on_zoom_change(self, _event=None) -> None:
        value = self.zoom_var.get()
        if value not in self._zoom_levels:
            return
        self.zoom_level = value
        self._zoom_factor = self._parse_zoom_level(value)
        self._update_named_fonts()
        self.setup_styles()
        self._apply_advanced_visibility()
        self._save_ui_preferences()

    def _on_toggle_advanced(self) -> None:
        self.show_advanced = bool(self.advanced_var.get())
        self._apply_advanced_visibility()
        self._save_ui_preferences()

    def _apply_advanced_visibility(self) -> None:
        if hasattr(self, "advanced_var"):
            self.advanced_var.set(bool(self.show_advanced))
        for widget, grid_info in self.advanced_cards:
            if self.show_advanced:
                widget.grid(**grid_info)
            else:
                widget.grid_remove()

    def _mark_advanced_card(self, widget: ttk.Widget) -> None:
        info = widget.grid_info().copy()
        info.pop("in", None)
        self.advanced_cards.append((widget, info))

    def _save_ui_preferences(self) -> None:
        self.ui_preferences["compact_mode"] = bool(self.compact_mode)
        self.ui_preferences["zoom_level"] = self.zoom_level
        self.ui_preferences["sidebar_collapsed"] = bool(self.sidebar_collapsed)
        self.ui_preferences["show_advanced"] = bool(self.show_advanced)
        save_settings(self.settings)

    def _persist_view_preferences(self) -> None:
        self.ui_preferences["window_geometry"] = self.geometry()
        self._save_ui_preferences()

    def log(self, message: str) -> None:
        """Append a message to the on-screen log in a thread-safe way."""

        if message is None:
            return

        def append() -> None:
            text = str(message)
            if hasattr(self, "log_area"):
                self.log_area.config(state=tk.NORMAL)
                self.log_area.insert(tk.END, text + "\n")
                self.log_area.see(tk.END)
                self.log_area.config(state=tk.DISABLED)
            else:
                print(text)

        if threading.current_thread() is threading.main_thread():
            append()
        else:
            self.after(0, append)

    def run_in_thread(self, target: Callable, *args, daemon: bool = True, **kwargs) -> threading.Thread:
        """Execute a callable in a background thread and report unexpected errors."""

        def worker() -> None:
            try:
                target(*args, **kwargs)
            except Exception as exc:  # pragma: no cover - defensive
                def report() -> None:
                    error_text = f"{self.tr('Error')}: {exc}"
                    self.log(error_text)
                    messagebox.showerror(self.tr("Error"), str(exc))

                self.after(0, report)

        thread = threading.Thread(target=worker, daemon=daemon)
        thread.start()
        return thread

    def task_completion_popup(self, status: str, message: str) -> None:
        """Display a completion dialog from background tasks on the main thread."""

        def show_dialog() -> None:
            normalized = (status or "").strip().lower()
            title = self.tr(status) if status else ""
            if normalized == "error":
                messagebox.showerror(self.tr("Error"), message)
            elif normalized == "warning":
                messagebox.showwarning(self.tr("Warning"), message)
            else:
                dialog_title = title or self.tr("Information")
                messagebox.showinfo(dialog_title, message)

        if threading.current_thread() is threading.main_thread():
            show_dialog()
        else:
            self.after(0, show_dialog)

    def setup_styles(self):
        """Configure a modern dark theme for the application widgets."""
        base_bg = "#0b1120"
        card_bg = "#111c2e"
        panel_header_bg = "#1a253a"
        panel_header_hover = "#24324d"
        panel_border = "#1f2d47"
        accent = "#38bdf8"
        accent_hover = "#0ea5e9"
        text_primary = "#f1f5f9"
        text_secondary = "#cbd5f5"
        text_muted = "#94a3b8"
        tooltip_bg = "#111c2e"
        tooltip_fg = text_primary
        sidebar_bg = "#081021"

        self.theme_colors = {
            "base_bg": base_bg,
            "card_bg": card_bg,
            "panel_header_bg": panel_header_bg,
            "panel_header_hover": panel_header_hover,
            "panel_border": panel_border,
            "accent": accent,
            "accent_hover": accent_hover,
            "text_primary": text_primary,
            "text_secondary": text_secondary,
            "text_muted": text_muted,
            "tooltip_bg": tooltip_bg,
            "tooltip_fg": tooltip_fg,
        }

        self.configure(bg=base_bg)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=base_bg)
        style.configure("Header.TFrame", background=base_bg)
        style.configure(
            "Card.TLabelframe",
            background=card_bg,
            borderwidth=0,
            padding=self._pad_value(15, 6),
        )
        style.configure("PanelBody.TFrame", background=card_bg)
        style.configure(
            "Card.TLabelframe.Label",
            background=card_bg,
            foreground=text_primary,
            font=self._font("Segoe UI Semibold", 11),
        )
        style.configure(
            "TLabel",
            background=card_bg,
            foreground=text_primary,
            font=self._font("Segoe UI", 10),
        )
        style.configure(
            "Description.TLabel",
            background=card_bg,
            foreground=text_muted,
            font=self._font("Segoe UI", 10),
        )
        style.configure(
            "Primary.TLabel",
            background=base_bg,
            foreground=text_primary,
            font=self._font("Segoe UI Semibold", 18),
        )
        style.configure(
            "Secondary.TLabel",
            background=base_bg,
            foreground=text_muted,
            font=self._font("Segoe UI", 11),
        )
        style.configure("Toolbar.TFrame", background=base_bg)
        style.configure(
            "TNotebook",
            background=base_bg,
            borderwidth=0,
            tabmargins=(4, 2, 4, 0),
        )
        style.configure(
            "TNotebook.Tab",
            background=card_bg,
            foreground=text_primary,
            padding=(self._pad_value(16, 6), self._pad_value(8, 3)),
            font=self._font("Segoe UI", 10),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", accent), ("active", accent_hover)],
            foreground=[("selected", text_primary), ("active", text_primary)],
        )
        style.configure(
            "TButton",
            background=accent,
            foreground=text_primary,
            font=self._font("Segoe UI Semibold", 10),
            padding=(self._pad_value(14, 6), self._pad_value(6, 3)),
            borderwidth=0,
        )
        style.map(
            "TButton",
            background=[("active", accent_hover), ("disabled", "#1e293b")],
            foreground=[("disabled", text_muted), ("active", text_primary)],
        )
        style.configure(
            "TEntry",
            fieldbackground="#111827",
            foreground=text_primary,
            insertcolor=text_primary,
            padding=self._pad_value(8, 4),
        )
        style.configure(
            "TCombobox",
            fieldbackground="#111827",
            foreground=text_primary,
            background=card_bg,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#111827"), ("disabled", "#1f2937")],
            foreground=[("readonly", text_primary), ("disabled", text_muted)],
        )
        style.configure(
            "Light.TCombobox",
            fieldbackground=card_bg,
            foreground=text_primary,
            background=card_bg,
        )
        style.map(
            "Light.TCombobox",
            fieldbackground=[("readonly", card_bg), ("disabled", "#1f2937")],
            foreground=[("readonly", text_primary), ("disabled", text_muted)],
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
            font=self._font("Segoe UI", 10),
        )
        style.configure(
            "TCheckbutton",
            background=base_bg,
            foreground=text_primary,
            font=self._font("Segoe UI", 10),
        )
        style.map(
            "TEntry",
            fieldbackground=[("disabled", "#1f2937")],
            foreground=[("disabled", text_muted)],
        )
        style.map(
            "TCheckbutton",
            background=[("active", card_bg)],
            foreground=[("disabled", text_muted)],
        )
        style.configure(
            "Sidebar.TFrame",
            background=sidebar_bg,
        )
        style.configure(
            "SidebarHeader.TFrame",
            background=sidebar_bg,
        )
        sidebar_padding = (self._pad_value(12, 6), self._pad_value(6, 3))
        style.configure(
            "Sidebar.TButton",
            background=card_bg,
            foreground=text_primary,
            font=self._font("Segoe UI", 10),
            padding=sidebar_padding,
        )
        style.map(
            "Sidebar.TButton",
            background=[("active", panel_header_hover)],
        )
        style.configure(
            "SidebarSelected.TButton",
            background=accent,
            foreground=text_primary,
            font=self._font("Segoe UI Semibold", 10),
            padding=sidebar_padding,
        )
        style.map(
            "SidebarSelected.TButton",
            background=[("active", accent_hover)],
            foreground=[("active", text_primary)],
        )
        style.configure("Horizontal.TSeparator", background="#1f2937")

        option_font = self._font("Segoe UI", 10)
        self.option_add("*TCombobox*Listbox.font", option_font)
        self.option_add("*TCombobox*Listbox.foreground", text_primary)
        self.option_add("*TCombobox*Listbox.background", card_bg)
        self.option_add("*Background", base_bg)
        self.option_add("*Entry.background", "#111827")
        self.option_add("*Entry.foreground", text_primary)
        self.option_add("*Listbox.background", card_bg)
        self.option_add("*Listbox.foreground", text_primary)
        self.option_add("*Font", option_font)
        self.option_add("*Foreground", text_primary)

        self._apply_log_theme()
        if hasattr(self, "help_text_area"):
            self.help_text_area.configure(font=("Helvetica", self._scaled_size(10)))
        self._update_nav_highlight()
        self._update_sidebar_toggle_text()

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

    def create_view_toolbar(self):
        toolbar = ttk.Frame(self, style="Toolbar.TFrame")
        toolbar.pack(fill="x", padx=10, pady=(6, 4))
        self.view_toolbar = toolbar

        left = ttk.Frame(toolbar, style="Toolbar.TFrame")
        left.pack(side="left", fill="x", expand=True)
        right = ttk.Frame(toolbar, style="Toolbar.TFrame")
        right.pack(side="right")
        self.view_toolbar_left = left
        self.view_toolbar_right = right

        self.sidebar_toggle = ttk.Button(left, command=self._toggle_sidebar, style="TButton")
        self.sidebar_toggle.pack(side="left")

        self.compact_var = tk.BooleanVar(value=self.compact_mode)
        compact_check = ttk.Checkbutton(
            left,
            text=self.tr("Compact Mode"),
            variable=self.compact_var,
            command=self._on_toggle_compact,
        )
        compact_check.pack(side="left", padx=(12, 0))
        self.register_widget(compact_check, "Compact Mode")

        zoom_label = ttk.Label(left, text=self.tr("Zoom"), style="Secondary.TLabel")
        zoom_label.pack(side="left", padx=(12, 6))
        self.register_widget(zoom_label, "Zoom")
        self.zoom_label = zoom_label

        self.zoom_var = tk.StringVar(value=self.zoom_level)
        zoom_box = ttk.Combobox(
            left,
            textvariable=self.zoom_var,
            values=self._zoom_levels,
            state="readonly",
            width=6,
        )
        zoom_box.pack(side="left")
        zoom_box.bind("<<ComboboxSelected>>", self._on_zoom_change)
        self.zoom_selector = zoom_box

        self.advanced_var = tk.BooleanVar(value=self.show_advanced)
        advanced_check = ttk.Checkbutton(
            left,
            text=self.tr("Advanced Settings"),
            variable=self.advanced_var,
            command=self._on_toggle_advanced,
        )
        advanced_check.pack(side="left", padx=(12, 0))
        self.register_widget(advanced_check, "Advanced Settings")

        self._update_sidebar_toggle_text()

    def create_language_selector(self):
        """Create the language selection combobox and bind change events."""
        parent = getattr(self, "view_toolbar_right", self)
        container = ttk.Frame(parent, style="Toolbar.TFrame")
        container.pack(side="right")

        label = ttk.Label(container, text=self.tr("Language"), style="Secondary.TLabel")
        label.pack(side="left", padx=(0, 8))
        self.register_widget(label, "Language")

        self.language_selector = ttk.Combobox(
            container,
            textvariable=self.language_var,
            state="readonly",
            width=14,
        )
        self.language_selector.pack(side="left")
        self.language_selector.bind("<<ComboboxSelected>>", self._on_language_change)

        self._refresh_language_options()

    def create_section_card(self, parent: ttk.Frame, title_key: str) -> ttk.Labelframe:
        """Create a labeled card container for a tool section."""

        card = ttk.Labelframe(parent, text=self.tr(title_key), style="Card.TLabelframe")
        self.register_widget(card, title_key)

        info_text = self.tr_info(title_key)
        if info_text:
            description = ttk.Label(
                card,
                text=info_text,
                style="Description.TLabel",
                wraplength=420,
                justify="left",
            )
            description.pack(fill="x", pady=(0, 12))
            self.section_descriptions.append((description, title_key))

        body = ttk.Frame(card, style="PanelBody.TFrame")
        body.pack(fill="both", expand=True)
        card.body = body
        return card

    def create_file_image_panels(self, parent: ttk.Frame):
        """Build cards for file and image related tools."""

        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        # Allow additional rows to expand if needed.
        parent.rowconfigure(2, weight=1)

        copy_card = self.create_section_card(parent, "1. Copy/Move Files by List")
        copy_card.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        copy_frame = copy_card.body
        copy_frame.columnconfigure(1, weight=1)

        self.source_folder = tk.StringVar(value=self.settings.get("source_folder", ""))
        self.target_folder = tk.StringVar(value=self.settings.get("target_folder", ""))
        self.numbers_file = tk.StringVar()

        src_label = ttk.Label(copy_frame, text=self.tr("Source Folder:"))
        src_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(src_label, "Source Folder:")
        ttk.Entry(copy_frame, textvariable=self.source_folder).grid(row=0, column=1, sticky="we", padx=6, pady=6)
        src_browse = ttk.Button(
            copy_frame,
            text=self.tr("Browse..."),
            command=lambda: self.source_folder.set(filedialog.askdirectory()),
        )
        src_browse.grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(src_browse, "Browse...")

        tgt_label = ttk.Label(copy_frame, text=self.tr("Target Folder:"))
        tgt_label.grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(tgt_label, "Target Folder:")
        ttk.Entry(copy_frame, textvariable=self.target_folder).grid(row=1, column=1, sticky="we", padx=6, pady=6)
        tgt_browse = ttk.Button(
            copy_frame,
            text=self.tr("Browse..."),
            command=lambda: self.target_folder.set(filedialog.askdirectory()),
        )
        tgt_browse.grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(tgt_browse, "Browse...")

        list_label = ttk.Label(copy_frame, text=self.tr("Numbers File (List):"))
        list_label.grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(list_label, "Numbers File (List):")
        ttk.Entry(copy_frame, textvariable=self.numbers_file).grid(row=2, column=1, sticky="we", padx=6, pady=6)
        list_browse = ttk.Button(
            copy_frame,
            text=self.tr("Browse..."),
            command=lambda: self.numbers_file.set(
                filedialog.askopenfilename(
                    filetypes=[
                        ("Excel", "*.xlsx *.xls"),
                        ("CSV/TXT", "*.csv *.txt"),
                        ("All Files", "*.*"),
                    ]
                )
            ),
        )
        list_browse.grid(row=2, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(list_browse, "Browse...")

        button_frame = ttk.Frame(copy_frame, style="PanelBody.TFrame")
        button_frame.grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=(4, 6))

        copy_button = ttk.Button(button_frame, text=self.tr("Copy Files"), command=lambda: self.start_process_files("copy"))
        copy_button.pack(side="left")
        self.register_widget(copy_button, "Copy Files")

        move_button = ttk.Button(button_frame, text=self.tr("Move Files"), command=lambda: self.start_process_files("move"))
        move_button.pack(side="left", padx=(8, 0))
        self.register_widget(move_button, "Move Files")

        save_button = ttk.Button(button_frame, text=self.tr("Save Settings"), command=self.save_folder_settings)
        save_button._text_icon_prefix = "⚙"
        save_button.pack(side="left", padx=(8, 0))
        self.register_widget(save_button, "Save Settings")

        heic_card = self.create_section_card(parent, "2. Convert HEIC to JPG")
        heic_card.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        heic_frame = heic_card.body
        heic_frame.columnconfigure(1, weight=1)

        self.heic_folder = tk.StringVar()
        heic_label = ttk.Label(heic_frame, text=self.tr("Folder with HEIC files:"))
        heic_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(heic_label, "Folder with HEIC files:")
        ttk.Entry(heic_frame, textvariable=self.heic_folder).grid(row=0, column=1, sticky="we", padx=6, pady=6)
        heic_browse = ttk.Button(
            heic_frame,
            text=self.tr("Browse..."),
            command=lambda: self.heic_folder.set(filedialog.askdirectory()),
        )
        heic_browse.grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(heic_browse, "Browse...")

        heic_button = ttk.Button(heic_frame, text=self.tr("Convert"), command=self.start_heic_conversion)
        heic_button.grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))
        self.register_widget(heic_button, "Convert")

        resize_card = self.create_section_card(parent, "3. Batch Image Resizer")
        resize_card.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        resize_frame = resize_card.body
        resize_frame.columnconfigure(1, weight=1)

        self.resize_folder = tk.StringVar()
        folder_label = ttk.Label(resize_frame, text=self.tr("Image Folder:"))
        folder_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(folder_label, "Image Folder:")
        ttk.Entry(resize_frame, textvariable=self.resize_folder).grid(row=0, column=1, sticky="we", padx=6, pady=6)
        folder_browse = ttk.Button(
            resize_frame,
            text=self.tr("Browse..."),
            command=lambda: self.resize_folder.set(filedialog.askdirectory()),
        )
        folder_browse.grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(folder_browse, "Browse...")

        mode_label = ttk.Label(resize_frame, text=self.tr("Resize Mode:"))
        mode_label.grid(row=1, column=0, sticky="w", padx=6, pady=(6, 2))
        self.register_widget(mode_label, "Resize Mode:")

        self.resize_mode = tk.StringVar(value="width")
        mode_frame = ttk.Frame(resize_frame, style="PanelBody.TFrame")
        mode_frame.grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=(6, 2))

        width_radio = ttk.Radiobutton(
            mode_frame,
            text=self.tr("By Width"),
            value="width",
            variable=self.resize_mode,
            command=self._update_resize_inputs,
        )
        width_radio.pack(side="left")
        self.register_widget(width_radio, "By Width", attr="text")

        percent_radio = ttk.Radiobutton(
            mode_frame,
            text=self.tr("By Percentage"),
            value="percentage",
            variable=self.resize_mode,
            command=self._update_resize_inputs,
        )
        percent_radio.pack(side="left", padx=(10, 0))
        self.register_widget(percent_radio, "By Percentage", attr="text")

        self.max_width = tk.StringVar(value="1920")
        self.resize_percentage = tk.StringVar(value="80")
        self.quality = tk.StringVar(value="85")

        width_label = ttk.Label(resize_frame, text=self.tr("Max Width:"))
        width_label.grid(row=2, column=0, sticky="w", padx=6, pady=2)
        self.register_widget(width_label, "Max Width:")
        self.max_width_entry = ttk.Entry(resize_frame, textvariable=self.max_width, width=10)
        self.max_width_entry.grid(row=2, column=1, sticky="w", padx=6, pady=2)

        percent_label = ttk.Label(resize_frame, text=self.tr("Percentage (%):"))
        percent_label.grid(row=3, column=0, sticky="w", padx=6, pady=2)
        self.register_widget(percent_label, "Percentage (%):")
        self.resize_percentage_entry = ttk.Entry(resize_frame, textvariable=self.resize_percentage, width=10)
        self.resize_percentage_entry.grid(row=3, column=1, sticky="w", padx=6, pady=2)

        quality_label = ttk.Label(resize_frame, text=self.tr("JPEG Quality (1-95):"))
        quality_label.grid(row=4, column=0, sticky="w", padx=6, pady=2)
        self.register_widget(quality_label, "JPEG Quality (1-95):")
        ttk.Entry(resize_frame, textvariable=self.quality, width=10).grid(row=4, column=1, sticky="w", padx=6, pady=2)

        resize_button = ttk.Button(resize_frame, text=self.tr("Resize & Compress"), command=self.start_resize_task)
        resize_button.grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=(6, 6))
        self.register_widget(resize_button, "Resize & Compress")

        self._update_resize_inputs()

    def create_view_in_room_tab(self, parent: ttk.Frame) -> None:
        """Create the View in Room preview tab."""

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        card = self.create_section_card(parent, "View in Room")
        card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        frame = card.body
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

        self.view_in_room_room_path = tk.StringVar()
        self.view_in_room_rug_path = tk.StringVar()

        room_label = ttk.Label(frame, text=self.tr("Room Image:"))
        room_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(room_label, "Room Image:")

        room_entry = ttk.Entry(frame, textvariable=self.view_in_room_room_path, state="readonly")
        room_entry.grid(row=0, column=1, sticky="we", padx=6, pady=6)

        room_button = ttk.Button(
            frame,
            text=self.tr("Browse..."),
            command=lambda: self._select_view_in_room_file("room"),
        )
        room_button.grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(room_button, "Browse...")

        rug_label = ttk.Label(frame, text=self.tr("Rug Image:"))
        rug_label.grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(rug_label, "Rug Image:")

        rug_entry = ttk.Entry(frame, textvariable=self.view_in_room_rug_path, state="readonly")
        rug_entry.grid(row=1, column=1, sticky="we", padx=6, pady=6)

        rug_button = ttk.Button(
            frame,
            text=self.tr("Browse..."),
            command=lambda: self._select_view_in_room_file("rug"),
        )
        rug_button.grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(rug_button, "Browse...")

        controls_label = ttk.Label(
            frame,
            text=self.tr("View in Room Controls"),
            wraplength=520,
            foreground="#555555",
        )
        controls_label.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(4, 8))
        self.register_widget(controls_label, "View in Room Controls")

        button_frame = ttk.Frame(frame, style="PanelBody.TFrame")
        button_frame.grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))

        preview_button = ttk.Button(
            button_frame,
            text=self.tr("Generate Preview"),
            command=lambda: self.generate_view_in_room_preview(reset_rug=True),
        )
        preview_button.pack(side="left")
        self.register_widget(preview_button, "Generate Preview")

        save_button = ttk.Button(button_frame, text=self.tr("Save Image"), command=self.save_view_in_room_image)
        save_button.pack(side="left", padx=(8, 0))
        self.register_widget(save_button, "Save Image")

        manual_button = ttk.Button(
            button_frame,
            text=self.tr("Manual Place Rug"),
            command=self._start_manual_rug_placement,
        )
        manual_button.pack(side="left", padx=(8, 0))
        self.view_in_room_manual_button = manual_button
        self.register_widget(manual_button, "Manual Place Rug")

        self.view_in_room_manual_prompt_var = tk.StringVar(value="")
        manual_prompt_label = ttk.Label(
            frame,
            textvariable=self.view_in_room_manual_prompt_var,
            foreground="#1d4ed8",
            wraplength=520,
        )
        manual_prompt_label.grid(row=4, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 2))
        manual_prompt_label.configure(anchor="w")
        self.view_in_room_manual_prompt_label = manual_prompt_label

        canvas_bg = self._get_canvas_background(frame)
        self.view_in_room_canvas = tk.Canvas(
            frame,
            width=self.view_in_room_canvas_size[0],
            height=self.view_in_room_canvas_size[1],
            highlightthickness=0,
            borderwidth=0,
            background=canvas_bg,
        )
        self.view_in_room_canvas.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=6, pady=(8, 6))
        self.view_in_room_canvas.bind("<Button-1>", self._on_view_in_room_canvas_left_click)
        self.view_in_room_canvas.bind("<B1-Motion>", self._on_view_in_room_canvas_left_drag)
        self.view_in_room_canvas.bind("<ButtonRelease-1>", self._on_view_in_room_canvas_left_release)
        self.view_in_room_canvas.bind("<Button-3>", self._on_view_in_room_canvas_right_click)
        self.view_in_room_canvas.bind("<B3-Motion>", self._on_view_in_room_canvas_right_drag)
        self.view_in_room_canvas.bind("<ButtonRelease-3>", self._on_view_in_room_canvas_right_release)
        self.view_in_room_canvas.bind("<MouseWheel>", self._on_view_in_room_mouse_wheel)
        self.view_in_room_canvas.bind("<Button-4>", lambda e: self._on_view_in_room_mouse_wheel(e, delta=120))
        self.view_in_room_canvas.bind("<Button-5>", lambda e: self._on_view_in_room_mouse_wheel(e, delta=-120))
        self.view_in_room_canvas.bind("<Motion>", self._on_view_in_room_mouse_motion)

        self.view_in_room_rug_selected = False
        self.view_in_room_controls_hidden = True
        self.view_in_room_control_icon_size = 26
        self.view_in_room_control_items = {"rotate": (), "scale": ()}
        self.view_in_room_icon_bounds = {}
        self.view_in_room_hide_icons_job = None

        self.view_in_room_preview_has_image = False
        self._show_view_in_room_message(self.tr("Preview will appear here."))

    def _get_canvas_background(self, widget: tk.Widget) -> str:
        try:
            return widget.cget("background")
        except tk.TclError:
            try:
                return widget.winfo_toplevel().cget("background")
            except tk.TclError:
                return "#f0f0f0"

    def _show_view_in_room_message(self, message: str) -> None:
        if not hasattr(self, "view_in_room_canvas"):
            return
        width, height = self.view_in_room_canvas_size
        self.view_in_room_canvas.delete("all")
        self.view_in_room_canvas.config(width=width, height=height)
        text_width = min(520, max(200, width - 40))
        self.view_in_room_canvas_message = self.view_in_room_canvas.create_text(
            width // 2,
            height // 2,
            text=message,
            fill="#666666",
            width=text_width,
            justify="center",
        )
        self.view_in_room_preview_has_image = False
        self.view_in_room_canvas_room_item = None
        self.view_in_room_canvas_rug_item = None
        self.view_in_room_canvas_room_photo = None
        self.view_in_room_canvas_rug_photo = None
        self.view_in_room_rug_display_bbox = None
        self.view_in_room_rug_display_center = None
        self._clear_view_in_room_selection()
        self.view_in_room_preview_image = None
        self.view_in_room_drag_mode = None
        self.view_in_room_drag_offset = (0.0, 0.0)
        self._reset_manual_placement_state()

    def _reset_manual_placement_state(self, *, reset_polygon: bool = True) -> None:
        self.view_in_room_manual_active = False
        self.view_in_room_manual_points = []
        if reset_polygon:
            self.view_in_room_manual_relative_polygon = None
        if hasattr(self, "view_in_room_canvas"):
            try:
                self.view_in_room_canvas.delete("manual-overlay")
            except tk.TclError:
                pass
        self._update_manual_prompt_label()

    def _update_manual_prompt_label(self) -> None:
        if self.view_in_room_manual_prompt_var is None:
            return
        if self.view_in_room_manual_active:
            index = len(self.view_in_room_manual_points)
            if index < 4:
                key = f"Manual Prompt {index + 1}"
                self.view_in_room_manual_prompt_var.set(self.tr(key))
            else:
                self.view_in_room_manual_prompt_var.set("")
        elif self.view_in_room_manual_relative_polygon:
            self.view_in_room_manual_prompt_var.set(self.tr("Manual Placement Complete"))
        else:
            self.view_in_room_manual_prompt_var.set("")

    def _start_manual_rug_placement(self) -> None:
        if not getattr(self, "view_in_room_preview_has_image", False):
            if self.view_in_room_manual_prompt_var is not None:
                self.view_in_room_manual_prompt_var.set("Lütfen oda ve halı görsellerini seçin.")
            return
        self.view_in_room_manual_active = True
        self.view_in_room_manual_points = []
        self._set_view_in_room_rug_selected(False)
        self._hide_view_in_room_control_icons()
        self._update_manual_prompt_label()
        self._render_view_in_room_canvas()

    def _handle_manual_placement_click(self, event: tk.Event) -> None:
        if not self.view_in_room_manual_active:
            return
        scale = self.view_in_room_display_scale or 1.0
        if scale <= 0:
            scale = 1.0
        actual_x = event.x / scale
        actual_y = event.y / scale
        self.view_in_room_manual_points.append((actual_x, actual_y))
        if len(self.view_in_room_manual_points) > 4:
            self.view_in_room_manual_points = self.view_in_room_manual_points[:4]
        if len(self.view_in_room_manual_points) == 4:
            self._finalize_manual_placement()
        else:
            self._update_manual_prompt_label()
            self._render_view_in_room_canvas()

    def _finalize_manual_placement(self) -> None:
        points = self.view_in_room_manual_points
        if len(points) != 4:
            return
        if self.view_in_room_rug_scale <= 0:
            self.view_in_room_rug_scale = 1.0
        center_x = sum(x for x, _ in points) / 4.0
        center_y = sum(y for _, y in points) / 4.0
        base_scale = self.view_in_room_rug_scale or 1.0
        relative = [((x - center_x) / base_scale, (y - center_y) / base_scale) for x, y in points]
        self.view_in_room_manual_relative_polygon = relative
        self.view_in_room_manual_active = False
        self.view_in_room_manual_points = []
        self.view_in_room_rug_angle = 0.0
        self.view_in_room_rug_center = (center_x, center_y)
        rug_size = self._get_current_rug_size()
        if self.view_in_room_rug_center is not None:
            self.view_in_room_rug_center = self._clamp_rug_center(self.view_in_room_rug_center, rug_size)
        self._set_view_in_room_rug_selected(True)
        self._update_manual_prompt_label()
        self._render_view_in_room_canvas()
        self._update_view_in_room_preview_image()

    def _draw_manual_overlay(self, room_scale: float) -> None:
        if not hasattr(self, "view_in_room_canvas"):
            return
        canvas = self.view_in_room_canvas
        try:
            canvas.delete("manual-overlay")
        except tk.TclError:
            return
        if self.view_in_room_manual_active:
            if self.view_in_room_manual_points:
                display_points = [
                    (x * room_scale, y * room_scale) for x, y in self.view_in_room_manual_points
                ]
                color = "#2563eb"
                if len(display_points) == 1:
                    x, y = display_points[0]
                    radius = 4
                    canvas.create_oval(
                        x - radius,
                        y - radius,
                        x + radius,
                        y + radius,
                        outline=color,
                        fill=color,
                        width=1,
                        tags="manual-overlay",
                        state="disabled",
                    )
                else:
                    flat = [coord for point in display_points for coord in point]
                    canvas.create_line(
                        *flat,
                        fill=color,
                        width=2,
                        tags="manual-overlay",
                        state="disabled",
                    )
                    if len(display_points) == 4:
                        canvas.create_polygon(
                            *flat,
                            outline=color,
                            fill=color,
                            stipple="gray25",
                            width=2,
                            tags="manual-overlay",
                            state="disabled",
                        )
                canvas.tag_raise("manual-overlay")
            return
        if (
            self.view_in_room_manual_relative_polygon
            and self.view_in_room_rug_center is not None
            and room_scale > 0
        ):
            angle_rad = math.radians(self.view_in_room_rug_angle % 360)
            cos_a = math.cos(angle_rad)
            sin_a = math.sin(angle_rad)
            points = []
            for dx, dy in self.view_in_room_manual_relative_polygon:
                scaled_x = dx * self.view_in_room_rug_scale
                scaled_y = dy * self.view_in_room_rug_scale
                rot_x = scaled_x * cos_a - scaled_y * sin_a
                rot_y = scaled_x * sin_a + scaled_y * cos_a
                actual_x = (self.view_in_room_rug_center[0] + rot_x) * room_scale
                actual_y = (self.view_in_room_rug_center[1] + rot_y) * room_scale
                points.extend([actual_x, actual_y])
            if points:
                canvas.create_polygon(
                    *points,
                    outline="#2563eb",
                    fill="#2563eb",
                    stipple="gray25",
                    width=2,
                    tags="manual-overlay",
                    state="disabled",
                )
                canvas.tag_raise("manual-overlay")

    def _compute_pil_perspective_coeffs(
        self, src: List[Tuple[float, float]], dst: List[Tuple[float, float]]
    ) -> Optional[List[float]]:
        matrix = []
        vector = []
        for (x_src, y_src), (x_dst, y_dst) in zip(src, dst):
            matrix.append([x_src, y_src, 1, 0, 0, 0, -x_dst * x_src, -x_dst * y_src])
            matrix.append([0, 0, 0, x_src, y_src, 1, -y_dst * x_src, -y_dst * y_src])
            vector.extend([x_dst, y_dst])
        matrix_np = np.array(matrix, dtype=float)
        vector_np = np.array(vector, dtype=float)
        try:
            solution = np.linalg.solve(matrix_np, vector_np)
        except np.linalg.LinAlgError:
            return None
        h11, h12, h13, h21, h22, h23, h31, h32 = solution
        transform_matrix = np.array(
            [[h11, h12, h13], [h21, h22, h23], [h31, h32, 1.0]], dtype=float
        )
        try:
            inverse = np.linalg.inv(transform_matrix)
        except np.linalg.LinAlgError:
            return None
        a, b, c = inverse[0]
        d, e, f = inverse[1]
        g, h, _ = inverse[2]
        return [a, b, c, d, e, f, g, h]

    def _create_warped_rug(
        self,
        scale_multiplier: float,
        *,
        angle: Optional[float] = None,
        source_image: Optional[Image.Image] = None,
    ) -> Optional[RugWarpResult]:
        base = source_image if source_image is not None else self.view_in_room_rug_original
        if base is None:
            return None
        resampling = getattr(Image, "Resampling", Image).BICUBIC
        width = max(1, int(round(base.width * scale_multiplier)))
        height = max(1, int(round(base.height * scale_multiplier)))
        if width <= 0 or height <= 0:
            return None
        rug_scaled = base.resize((width, height), resample=resampling)
        manual_relative = self.view_in_room_manual_relative_polygon
        if manual_relative and len(manual_relative) == 4:
            dest_local = [
                (dx * scale_multiplier, dy * scale_multiplier) for dx, dy in manual_relative
            ]
        else:
            top_scale = getattr(self, "view_in_room_perspective_top_scale", 0.6)
            top_scale = float(max(0.1, min(top_scale, 0.95)))
            bottom_half_width = width / 2.0
            top_half_width = bottom_half_width * top_scale
            half_height = height / 2.0
            dest_local = [
                (-top_half_width, -half_height),
                (top_half_width, -half_height),
                (bottom_half_width, half_height),
                (-bottom_half_width, half_height),
            ]
        angle_value = angle if angle is not None else self.view_in_room_rug_angle
        angle_rad = math.radians(angle_value % 360)
        cos_a = math.cos(angle_rad)
        sin_a = math.sin(angle_rad)
        dest_rotated = [
            (x * cos_a - y * sin_a, x * sin_a + y * cos_a)
            for x, y in dest_local
        ]
        xs = [pt[0] for pt in dest_rotated]
        ys = [pt[1] for pt in dest_rotated]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        out_width = int(math.ceil(max_x - min_x))
        out_height = int(math.ceil(max_y - min_y))
        if out_width <= 0 or out_height <= 0:
            return None
        dest_shifted = [(x - min_x, y - min_y) for x, y in dest_rotated]
        src = [(0.0, 0.0), (float(width), 0.0), (float(width), float(height)), (0.0, float(height))]
        coeffs = self._compute_pil_perspective_coeffs(src, dest_shifted)
        if coeffs is None:
            return None
        warped = rug_scaled.transform(
            (out_width, out_height), Image.PERSPECTIVE, coeffs, resample=resampling
        )
        offset = (min_x, min_y)
        return RugWarpResult(warped, offset, (out_width, out_height), dest_rotated)

    def _default_rug_center(self, room_img: Image.Image) -> Tuple[float, float]:
        return (room_img.width / 2.0, room_img.height * 0.75)

    def _calculate_default_rug_scale(self, room_img: Image.Image, rug_img: Image.Image) -> float:
        if rug_img.width == 0 or rug_img.height == 0:
            return 1.0
        projection = self._create_warped_rug(1.0, angle=0.0, source_image=rug_img)
        if projection is None:
            proj_width, proj_height = rug_img.width, rug_img.height
        else:
            proj_width, proj_height = projection.size
        target_width = room_img.width * 0.45
        target_height = room_img.height * 0.3
        scale_width = target_width / proj_width
        scale_height = target_height / proj_height
        scale = min(scale_width, scale_height)
        return float(max(0.2, min(scale, 1.6)))

    def _get_transformed_rug(self, scale_multiplier: float) -> Optional[RugWarpResult]:
        return self._create_warped_rug(scale_multiplier)

    def _get_current_rug_size(self) -> Tuple[int, int]:
        transformed = self._get_transformed_rug(self.view_in_room_rug_scale)
        if transformed is None:
            return 0, 0
        return transformed.size

    def _clamp_rug_center(self, center: Tuple[float, float], rug_size: Tuple[int, int]) -> Tuple[float, float]:
        room_img = self.view_in_room_room_image
        if room_img is None:
            return center
        width, height = rug_size
        half_w = width / 2.0
        half_h = height / 2.0
        max_x = max(half_w, room_img.width - half_w)
        max_y = max(half_h, room_img.height - half_h)
        clamped_x = min(max(center[0], half_w), max_x)
        clamped_y = min(max(center[1], half_h), max_y)
        return (clamped_x, clamped_y)

    def _clear_view_in_room_selection(self) -> None:
        self.view_in_room_rug_selected = False
        self._hide_view_in_room_control_icons()

    def _set_view_in_room_rug_selected(self, selected: bool) -> None:
        if selected:
            if not self.view_in_room_rug_selected:
                self.view_in_room_rug_selected = True
            self._show_view_in_room_control_icons()
        else:
            if self.view_in_room_rug_selected:
                self.view_in_room_rug_selected = False
            self._hide_view_in_room_control_icons()

    def _show_view_in_room_control_icons(self) -> None:
        if self.view_in_room_controls_hidden:
            self.view_in_room_controls_hidden = False
        self._ensure_view_in_room_control_icons()
        self._schedule_view_in_room_icon_hide()

    def _hide_view_in_room_control_icons(self) -> None:
        self.view_in_room_controls_hidden = True
        self._remove_view_in_room_control_icons()
        self._cancel_view_in_room_icon_hide()

    def _remove_view_in_room_control_icons(self) -> None:
        if not hasattr(self, "view_in_room_canvas"):
            return
        for items in self.view_in_room_control_items.values():
            for item in items or ():
                if item:
                    self.view_in_room_canvas.delete(item)
        self.view_in_room_control_items = {"rotate": (), "scale": ()}
        self.view_in_room_icon_bounds = {}

    def _ensure_view_in_room_control_icons(self) -> None:
        if (
            not getattr(self, "view_in_room_canvas", None)
            or not self.view_in_room_rug_selected
            or self.view_in_room_controls_hidden
        ):
            self._remove_view_in_room_control_icons()
            return
        if self.view_in_room_rug_display_bbox is None:
            self._remove_view_in_room_control_icons()
            return
        x0, y0, x1, y1 = self.view_in_room_rug_display_bbox
        icon_size = self.view_in_room_control_icon_size
        radius = icon_size / 2.0
        margin = max(6, radius)

        positions = {
            "rotate": (x0 + margin, y0 + margin, "↻"),
            "scale": (x1 - margin, y1 - margin, "⤡"),
        }

        for name, (cx, cy, symbol) in positions.items():
            x_start = cx - radius
            y_start = cy - radius
            x_end = cx + radius
            y_end = cy + radius
            items = self.view_in_room_control_items.get(name)
            if items and len(items) == 2 and all(items):
                bg_id, text_id = items
                self.view_in_room_canvas.coords(bg_id, x_start, y_start, x_end, y_end)
                self.view_in_room_canvas.coords(text_id, cx, cy)
            else:
                bg_id = self.view_in_room_canvas.create_oval(
                    x_start,
                    y_start,
                    x_end,
                    y_end,
                    fill="#ffffff",
                    outline="#4a4a4a",
                    width=1,
                )
                text_id = self.view_in_room_canvas.create_text(
                    cx,
                    cy,
                    text=symbol,
                    fill="#333333",
                    font=("Segoe UI", 10, "bold"),
                )
                self.view_in_room_control_items[name] = (bg_id, text_id)
            self.view_in_room_icon_bounds[name] = (x_start, y_start, x_end, y_end)
            self.view_in_room_canvas.tag_raise(self.view_in_room_control_items[name][0])
            self.view_in_room_canvas.tag_raise(self.view_in_room_control_items[name][1])

    def _schedule_view_in_room_icon_hide(self) -> None:
        if not getattr(self, "view_in_room_canvas", None):
            return
        self._cancel_view_in_room_icon_hide()
        if not self.view_in_room_rug_selected:
            return
        self.view_in_room_hide_icons_job = self.view_in_room_canvas.after(
            10000, self._hide_view_in_room_control_icons_due_to_inactivity
        )

    def _cancel_view_in_room_icon_hide(self) -> None:
        if self.view_in_room_hide_icons_job and getattr(self, "view_in_room_canvas", None):
            try:
                self.view_in_room_canvas.after_cancel(self.view_in_room_hide_icons_job)
            except tk.TclError:
                pass
        self.view_in_room_hide_icons_job = None

    def _hide_view_in_room_control_icons_due_to_inactivity(self) -> None:
        self.view_in_room_hide_icons_job = None
        if not self.view_in_room_rug_selected:
            return
        self.view_in_room_controls_hidden = True
        self._remove_view_in_room_control_icons()

    def _record_view_in_room_mouse_activity(self) -> None:
        if not self.view_in_room_rug_selected:
            return
        if self.view_in_room_controls_hidden:
            self.view_in_room_controls_hidden = False
        self._ensure_view_in_room_control_icons()
        self._schedule_view_in_room_icon_hide()

    def _view_in_room_icon_hit_test(self, name: str, x: float, y: float) -> bool:
        bounds = self.view_in_room_icon_bounds.get(name)
        if not bounds:
            return False
        x0, y0, x1, y1 = bounds
        return x0 <= x <= x1 and y0 <= y <= y1

    def _on_view_in_room_mouse_motion(self, _event: tk.Event) -> None:
        self._record_view_in_room_mouse_activity()

    def _render_view_in_room_canvas(self) -> None:
        room_img = self.view_in_room_room_image
        if room_img is None:
            self._show_view_in_room_message(self.tr("Preview will appear here."))
            return
        resampling = getattr(Image, "Resampling", Image).LANCZOS
        max_width, max_height = self.view_in_room_display_size
        if room_img.width <= 0 or room_img.height <= 0:
            self._show_view_in_room_message(self.tr("Preview will appear here."))
            return
        scale = min(max_width / room_img.width, max_height / room_img.height, 1.0)
        if scale <= 0:
            scale = 1.0
        display_width = max(1, int(round(room_img.width * scale)))
        display_height = max(1, int(round(room_img.height * scale)))
        self.view_in_room_display_scale = scale
        self.view_in_room_canvas_size = (display_width, display_height)
        self.view_in_room_canvas.config(width=display_width, height=display_height)
        room_display = (
            room_img.resize((display_width, display_height), resample=resampling)
            if scale != 1.0
            else room_img.copy()
        )
        self.view_in_room_canvas.delete("all")
        self.view_in_room_canvas_room_photo = ImageTk.PhotoImage(room_display)
        self.view_in_room_canvas_room_item = self.view_in_room_canvas.create_image(
            0,
            0,
            anchor="nw",
            image=self.view_in_room_canvas_room_photo,
        )
        if self.view_in_room_rug_original is None or self.view_in_room_rug_center is None:
            self.view_in_room_canvas_rug_item = None
            self.view_in_room_canvas_rug_photo = None
            self.view_in_room_rug_display_bbox = None
            self.view_in_room_rug_display_center = None
            self.view_in_room_preview_has_image = False
            self._clear_view_in_room_selection()
            self._draw_manual_overlay(scale)
            return
        display_scale = self.view_in_room_rug_scale * scale
        rug_projection = self._get_transformed_rug(display_scale)
        if rug_projection is None:
            self.view_in_room_canvas_rug_item = None
            self.view_in_room_canvas_rug_photo = None
            self.view_in_room_rug_display_bbox = None
            self.view_in_room_rug_display_center = None
            self.view_in_room_preview_has_image = False
            self._clear_view_in_room_selection()
            self._draw_manual_overlay(scale)
            return
        rug_display = rug_projection.image
        center_x = self.view_in_room_rug_center[0] * scale
        center_y = self.view_in_room_rug_center[1] * scale
        offset_x, offset_y = rug_projection.offset
        top_left_x = center_x + offset_x
        top_left_y = center_y + offset_y
        self.view_in_room_canvas_rug_photo = ImageTk.PhotoImage(rug_display)
        self.view_in_room_canvas_rug_item = self.view_in_room_canvas.create_image(
            int(round(top_left_x)),
            int(round(top_left_y)),
            anchor="nw",
            image=self.view_in_room_canvas_rug_photo,
        )
        self.view_in_room_canvas_message = None
        self.view_in_room_rug_display_bbox = (
            top_left_x,
            top_left_y,
            top_left_x + rug_projection.size[0],
            top_left_y + rug_projection.size[1],
        )
        self.view_in_room_rug_display_center = (center_x, center_y)
        self.view_in_room_preview_has_image = True
        self._update_view_in_room_preview_image()
        self._draw_manual_overlay(scale)
        self._ensure_view_in_room_control_icons()

    def _on_view_in_room_canvas_left_click(self, event: tk.Event) -> None:
        if self.view_in_room_manual_active:
            self._handle_manual_placement_click(event)
            return
        if not getattr(self, "view_in_room_preview_has_image", False):
            return
        if self.view_in_room_rug_display_bbox is None or self.view_in_room_rug_display_center is None:
            return
        if self._view_in_room_icon_hit_test("rotate", event.x, event.y):
            center_x, center_y = self.view_in_room_rug_display_center
            self.view_in_room_drag_mode = "rotate_icon"
            self.view_in_room_rotation_reference = math.degrees(
                math.atan2(event.y - center_y, event.x - center_x)
            )
            self.view_in_room_rotation_start_angle = self.view_in_room_rug_angle
            self._set_view_in_room_rug_selected(True)
            self._record_view_in_room_mouse_activity()
            return
        if self._view_in_room_icon_hit_test("scale", event.x, event.y):
            center_x, center_y = self.view_in_room_rug_display_center
            self.view_in_room_drag_mode = "scale"
            self.view_in_room_scale_reference_distance = math.hypot(
                event.x - center_x, event.y - center_y
            )
            self.view_in_room_scale_start_value = self.view_in_room_rug_scale
            self._set_view_in_room_rug_selected(True)
            self._record_view_in_room_mouse_activity()
            return
        x0, y0, x1, y1 = self.view_in_room_rug_display_bbox
        if x0 <= event.x <= x1 and y0 <= event.y <= y1:
            self.view_in_room_drag_mode = "move"
            center_x, center_y = self.view_in_room_rug_display_center
            self.view_in_room_drag_offset = (event.x - center_x, event.y - center_y)
            if self.view_in_room_canvas_rug_item is not None:
                self.view_in_room_canvas.tag_raise(self.view_in_room_canvas_rug_item)
            self._set_view_in_room_rug_selected(True)
            self._record_view_in_room_mouse_activity()
        else:
            self._set_view_in_room_rug_selected(False)

    def _on_view_in_room_canvas_left_drag(self, event: tk.Event) -> None:
        if self.view_in_room_manual_active:
            return
        if self.view_in_room_drag_mode not in {"move", "scale", "rotate_icon"}:
            return
        self._record_view_in_room_mouse_activity()
        if self.view_in_room_drag_mode == "move":
            scale = self.view_in_room_display_scale or 1.0
            offset_x, offset_y = self.view_in_room_drag_offset
            new_center_display = (event.x - offset_x, event.y - offset_y)
            new_center_actual = (new_center_display[0] / scale, new_center_display[1] / scale)
            rug_size = self._get_current_rug_size()
            self.view_in_room_rug_center = self._clamp_rug_center(new_center_actual, rug_size)
        elif self.view_in_room_drag_mode == "scale":
            if self.view_in_room_rug_display_center is None:
                return
            center_x, center_y = self.view_in_room_rug_display_center
            start_distance = self.view_in_room_scale_reference_distance
            if start_distance <= 0:
                return
            current_distance = math.hypot(event.x - center_x, event.y - center_y)
            if current_distance <= 0:
                return
            ratio = current_distance / start_distance
            new_scale = self.view_in_room_scale_start_value * ratio
            new_scale = max(0.2, min(new_scale, 3.0))
            self.view_in_room_rug_scale = new_scale
            rug_size = self._get_current_rug_size()
            if self.view_in_room_rug_center is not None:
                self.view_in_room_rug_center = self._clamp_rug_center(
                    self.view_in_room_rug_center, rug_size
                )
        elif self.view_in_room_drag_mode == "rotate_icon":
            if self.view_in_room_rug_display_center is None:
                return
            center_x, center_y = self.view_in_room_rug_display_center
            current_angle = math.degrees(math.atan2(event.y - center_y, event.x - center_x))
            delta = current_angle - self.view_in_room_rotation_reference
            self.view_in_room_rug_angle = (self.view_in_room_rotation_start_angle + delta) % 360
        self._render_view_in_room_canvas()

    def _on_view_in_room_canvas_left_release(self, _event: tk.Event) -> None:
        if self.view_in_room_manual_active:
            return
        if self.view_in_room_drag_mode == "move":
            self.view_in_room_drag_mode = None
            self.view_in_room_drag_offset = (0.0, 0.0)
            self._update_view_in_room_preview_image()
        elif self.view_in_room_drag_mode == "scale":
            self.view_in_room_drag_mode = None
            self.view_in_room_scale_reference_distance = 0.0
            self._update_view_in_room_preview_image()
        elif self.view_in_room_drag_mode == "rotate_icon":
            self.view_in_room_drag_mode = None
            self._update_view_in_room_preview_image()

    def _on_view_in_room_canvas_right_click(self, event: tk.Event) -> None:
        if self.view_in_room_manual_active:
            return
        if not getattr(self, "view_in_room_preview_has_image", False):
            return
        if self.view_in_room_rug_display_bbox is None or self.view_in_room_rug_display_center is None:
            return
        x0, y0, x1, y1 = self.view_in_room_rug_display_bbox
        if not (x0 <= event.x <= x1 and y0 <= event.y <= y1):
            self._set_view_in_room_rug_selected(False)
            return
        center_x, center_y = self.view_in_room_rug_display_center
        self.view_in_room_drag_mode = "rotate"
        self.view_in_room_rotation_reference = math.degrees(math.atan2(event.y - center_y, event.x - center_x))
        self.view_in_room_rotation_start_angle = self.view_in_room_rug_angle
        self._set_view_in_room_rug_selected(True)
        self._record_view_in_room_mouse_activity()

    def _on_view_in_room_canvas_right_drag(self, event: tk.Event) -> None:
        if self.view_in_room_manual_active:
            return
        if self.view_in_room_drag_mode != "rotate":
            return
        if self.view_in_room_rug_display_center is None:
            return
        center_x, center_y = self.view_in_room_rug_display_center
        current_angle = math.degrees(math.atan2(event.y - center_y, event.x - center_x))
        delta = current_angle - self.view_in_room_rotation_reference
        self.view_in_room_rug_angle = (self.view_in_room_rotation_start_angle + delta) % 360
        self._record_view_in_room_mouse_activity()
        self._render_view_in_room_canvas()

    def _on_view_in_room_canvas_right_release(self, _event: tk.Event) -> None:
        if self.view_in_room_manual_active:
            return
        if self.view_in_room_drag_mode == "rotate":
            self.view_in_room_drag_mode = None
            self._update_view_in_room_preview_image()

    def _on_view_in_room_mouse_wheel(self, event: tk.Event, delta: Optional[int] = None) -> None:
        if self.view_in_room_manual_active:
            return
        if not getattr(self, "view_in_room_preview_has_image", False):
            return
        self._record_view_in_room_mouse_activity()
        wheel_delta = delta if delta is not None else getattr(event, "delta", 0)
        if wheel_delta == 0:
            return
        factor = 1.0 + (0.08 if wheel_delta > 0 else -0.08)
        new_scale = self.view_in_room_rug_scale * factor
        new_scale = max(0.2, min(new_scale, 3.0))
        self.view_in_room_rug_scale = new_scale
        rug_size = self._get_current_rug_size()
        if self.view_in_room_rug_center is not None:
            self.view_in_room_rug_center = self._clamp_rug_center(self.view_in_room_rug_center, rug_size)
        self._render_view_in_room_canvas()

    def _update_view_in_room_preview_image(self) -> None:
        room_img = self.view_in_room_room_image
        if room_img is None or self.view_in_room_rug_original is None:
            self.view_in_room_preview_image = None
            self.view_in_room_preview_has_image = False
            return
        if self.view_in_room_rug_center is None:
            self.view_in_room_preview_image = None
            self.view_in_room_preview_has_image = False
            return
        rug_projection = self._get_transformed_rug(self.view_in_room_rug_scale)
        if rug_projection is None:
            self.view_in_room_preview_image = None
            self.view_in_room_preview_has_image = False
            return
        rug_image = rug_projection.image
        rug_size = rug_projection.size
        self.view_in_room_rug_center = self._clamp_rug_center(self.view_in_room_rug_center, rug_size)
        center_x, center_y = self.view_in_room_rug_center
        offset_x, offset_y = rug_projection.offset
        top_left_x = int(round(center_x + offset_x))
        top_left_y = int(round(center_y + offset_y))
        overlay = Image.new("RGBA", room_img.size, (0, 0, 0, 0))
        overlay.paste(rug_image, (top_left_x, top_left_y), rug_image)
        composed = Image.alpha_composite(room_img, overlay)
        self.view_in_room_preview_image = composed
        self.view_in_room_preview_has_image = True

    def _select_view_in_room_file(self, target: str) -> None:
        """Prompt the user for an image path and store it in the relevant variable."""

        filetypes = [
            (self.tr("Image Files"), "*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff"),
            ("All Files", "*.*"),
        ]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return
        if target == "room":
            self.view_in_room_room_path.set(path)
        else:
            self.view_in_room_rug_path.set(path)
        self.generate_view_in_room_preview(reset_rug=(target == "rug"))

    def generate_view_in_room_preview(self, reset_rug: bool = True) -> None:
        """Create an overlaid preview of the rug inside the room photo."""

        if not hasattr(self, "view_in_room_room_path"):
            return

        room_path = self.view_in_room_room_path.get()
        rug_path = self.view_in_room_rug_path.get()
        if not room_path:
            messagebox.showwarning(self.tr("Warning"), self.tr("Please select both room and rug images."))
            return

        try:
            with Image.open(room_path) as room_raw:
                room_img = room_raw.convert("RGBA")
        except Exception as exc:  # pragma: no cover - safeguard for Pillow errors
            messagebox.showerror(
                self.tr("Error"),
                self.tr("Could not open selected images: {error}").format(error=exc),
            )
            return

        self.view_in_room_room_image = room_img
        if reset_rug:
            self._reset_manual_placement_state()

        if not rug_path:
            if reset_rug:
                messagebox.showwarning(self.tr("Warning"), self.tr("Please select both room and rug images."))
            self.view_in_room_rug_original = None
            self.view_in_room_rug_center = None
            self.view_in_room_rug_scale = 1.0
            self.view_in_room_rug_angle = 0.0
            self.view_in_room_preview_has_image = False
            self._render_view_in_room_canvas()
            return

        try:
            with Image.open(rug_path) as rug_raw:
                rug_img = rug_raw.convert("RGBA")
        except Exception as exc:  # pragma: no cover - safeguard for Pillow errors
            messagebox.showerror(
                self.tr("Error"),
                self.tr("Could not open selected images: {error}").format(error=exc),
            )
            return

        self.view_in_room_rug_original = rug_img
        if reset_rug or self.view_in_room_rug_center is None:
            self.view_in_room_rug_scale = self._calculate_default_rug_scale(room_img, rug_img)
            self.view_in_room_rug_angle = 0.0
            self.view_in_room_rug_center = self._default_rug_center(room_img)
        rug_size = self._get_current_rug_size()
        if self.view_in_room_rug_center is not None:
            self.view_in_room_rug_center = self._clamp_rug_center(self.view_in_room_rug_center, rug_size)
        self.view_in_room_preview_photo = None
        self._render_view_in_room_canvas()
    def save_view_in_room_image(self) -> None:
        """Save the generated preview to disk."""

        if not getattr(self, "view_in_room_preview_image", None):
            messagebox.showwarning(
                self.tr("Warning"),
                self.tr("No preview available. Please generate a preview first."),
            )
            return

        filetypes = [
            ("PNG", "*.png"),
            ("JPEG", "*.jpg *.jpeg"),
            ("TIFF", "*.tif *.tiff"),
            ("All Files", "*.*"),
        ]
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=filetypes)
        if not path:
            return

        image_to_save = self.view_in_room_preview_image
        ext = os.path.splitext(path)[1].lower()
        if ext in {".jpg", ".jpeg"} and image_to_save.mode != "RGB":
            image_to_save = image_to_save.convert("RGB")

        try:
            image_to_save.save(path)
        except Exception as exc:  # pragma: no cover - filesystem errors
            messagebox.showerror(self.tr("Error"), self.tr("File could not be saved: {error}").format(error=exc))
            return

        messagebox.showinfo(self.tr("Success"), self.tr("Preview image saved to {path}.").format(path=path))

    def tr(self, text_key):
        """Translate a text key according to the selected language."""
        return translations.get(self.language, translations["en"]).get(text_key, text_key)

    def update_language(self, lang: str) -> None:
        """Update the UI language immediately without restarting the app."""
        if lang not in translations:
            return
        self.language = lang
        self.settings["language"] = lang
        save_settings(self.settings)
        self.refresh_translations()

    def tr_info(self, title_key: str) -> Optional[str]:
        """Return the localized info text for a given panel title."""
        language_info = PANEL_INFO.get(self.language) or PANEL_INFO.get("en", {})
        info_text = language_info.get(title_key)
        if info_text:
            return info_text
        return PANEL_INFO.get("en", {}).get(title_key)

    def _update_resize_inputs(self):
        """Enable the relevant resize input fields according to the selected mode."""
        if not hasattr(self, "resize_mode"):
            return
        max_width_entry = getattr(self, "max_width_entry", None)
        percentage_entry = getattr(self, "resize_percentage_entry", None)

        if max_width_entry is None or percentage_entry is None:
            return

        if self.resize_mode.get() == "width":
            max_width_entry.config(state="normal")
            percentage_entry.config(state="disabled")
        else:
            max_width_entry.config(state="disabled")
            percentage_entry.config(state="normal")

    def register_widget(self, widget, text_key, attr="text"):
        """Register a widget for translation updates."""
        self.translatable_widgets.append((widget, attr, text_key))
        self._apply_translation(widget, attr, text_key)

    def _apply_translation(self, widget, attr, text_key):
        try:
            value = self.tr(text_key)
            if attr == "text":
                prefix = getattr(widget, "_text_icon_prefix", "")
                suffix = getattr(widget, "_text_icon_suffix", "")
                if prefix:
                    value = f"{prefix} {value}"
                if suffix:
                    value = f"{value} {suffix}"
            widget.configure(**{attr: value})
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
        for label, title_key in self.section_descriptions:
            info_text = self.tr_info(title_key)
            if info_text:
                label.configure(text=info_text)
            else:
                label.configure(text="")
        if hasattr(self, "section_notebook"):
            for frame, title_key in self.notebook_tabs:
                self.section_notebook.tab(frame, text=self.tr(title_key))
        self.update_help_tab_content()
        if hasattr(self, "shared_status_var"):
            self._apply_shared_status_translation()
        if hasattr(self, "shared_help_text"):
            self._update_shared_help_text(self.shared_status_host, self.shared_status_port)
        if hasattr(self, "rug_manual_history_var"):
            self._update_manual_history_display()
        if hasattr(self, "rug_manual_result_var"):
            self._update_manual_result_label()
        if hasattr(self, "rug_result_tree"):
            self._refresh_rug_comparison_display()
        if hasattr(self, "rug_control_tree"):
            self.populate_rug_no_control_tree(getattr(self, "rug_control_results", []))
        if hasattr(self, "wayfair_formatter"):
            self.wayfair_formatter.set_translator(self.tr)
        if hasattr(self, "shared_printer_server"):
            self.shared_printer_server.set_translator(self.tr)
        if hasattr(self, "view_in_room_canvas") and not self.view_in_room_preview_has_image:
            self._show_view_in_room_message(self.tr("Preview will appear here."))
        self._update_manual_prompt_label()
        self._update_sidebar_toggle_text()
        self._refresh_language_options()

    def _refresh_language_options(self):
        """Update language combobox options according to current translations."""
        if not hasattr(self, "language_selector"):
            return
        values = [self.tr(key) for key in self.language_options.values()]
        self._updating_language_selector = True
        self.language_selector.configure(values=values)
        current_display = self.tr(self.language_options.get(self.language, "English"))
        self.language_var.set(current_display)
        self._updating_language_selector = False

    def _on_language_change(self, event=None):
        """Handle user selection of a different UI language."""
        if self._updating_language_selector:
            return
        selected = self.language_var.get()
        for code, key in self.language_options.items():
            if selected == self.tr(key):
                if code != self.language:
                    self.update_language(code)
                    self.log(
                        self.tr("Language changed to {language}.").format(
                            language=self.tr(key)
                        )
                    )
                else:
                    self.refresh_translations()
                break
        else:
            self._refresh_language_options()

    def update_help_tab_content(self):
        if hasattr(self, "help_text_area"):
            self.help_text_area.config(state=tk.NORMAL)
            self.help_text_area.delete("1.0", tk.END)
            help_content = self.tr("ABOUT_CONTENT").format(version=__version__)
            self.help_text_area.insert(tk.END, help_content)
            self.help_text_area.config(state=tk.DISABLED)

    def _format_shared_status(self, state: str, host: Optional[str], port: Optional[int]) -> str:
        """Durum anahtarına göre kullanıcıya gösterilecek metni hazırlar."""
        mapping = {
            "running": "Server Status: 🟢 Running",
            "stopped": "Server Status: ⚪ Stopped",
            "checking": "Server Status: ⏳ Checking...",
        }
        if state == "running" and host and port:
            return self.tr("SHARED_PRINTER_STATUS_DETAIL").format(host=host, port=port)
        return self.tr(mapping.get(state, "Server Status: ⚪ Stopped"))

    def _set_shared_status(self, state: str, host: Optional[str] = None, port: Optional[int] = None) -> None:
        """Durum değiştiğinde etiket ve yardım metnini günceller."""
        with self.shared_status_lock:
            self.shared_status_state = state
            if host is not None:
                self.shared_status_host = host
            if port is not None:
                self.shared_status_port = port
            host_value = self.shared_status_host or self.shared_printer_server.current_host()
            port_value = self.shared_status_port
            if port_value is None:
                current_port = self.shared_printer_server.current_port()
                if current_port:
                    port_value = current_port
            status_text = self._format_shared_status(state, host_value, port_value)
            self.shared_status_var.set(status_text)
        self._update_shared_help_text(host_value, port_value)

    def _apply_shared_status_translation(self) -> None:
        """Dil değiştiğinde mevcut durumu tekrar yazar."""
        with self.shared_status_lock:
            state = self.shared_status_state
            host = self.shared_status_host
            port = self.shared_status_port
        self.shared_status_var.set(self._format_shared_status(state, host, port))

    def _update_shared_help_text(self, host: Optional[str] = None, port: Optional[int] = None) -> None:
        """Yardım kutusundaki açıklama metnini günceller."""
        if not hasattr(self, "shared_help_text"):
            return
        host_value = host or self.shared_printer_server.current_host() or resolve_local_ip() or "127.0.0.1"
        port_value = port
        if port_value is None:
            try:
                port_value = int(self.shared_port_var.get().strip())
            except (ValueError, AttributeError):
                port_value = 5151
            current_port = self.shared_printer_server.current_port()
            if current_port:
                port_value = current_port
        message = self.tr("SHARED_PRINTER_HELP_TEXT").format(host=host_value, port=port_value)
        self.shared_help_text.config(state=tk.NORMAL)
        self.shared_help_text.delete("1.0", tk.END)
        self.shared_help_text.insert(tk.END, message)
        self.shared_help_text.config(state=tk.DISABLED)

    def start_shared_printer(self) -> None:
        """Gömülü Flask sunucusunu başlatır."""
        self._start_shared_printer(show_dialogs=True)

    def _start_shared_printer(self, *, show_dialogs: bool) -> bool:
        """Paylaşılan yazıcı sunucusunu başlatır ve başarı durumunu döndürür."""
        if self.shared_printer_server.is_running():
            host = self.shared_printer_server.current_host()
            port = self.shared_printer_server.current_port()
            self._set_shared_status("running", host, port)
            return True

        shared_settings = self.settings.setdefault("shared_label_printer", {})

        token = self.shared_token_var.get().strip()
        if not token:
            message = self.tr("SHARED_PRINTER_TOKEN_REQUIRED")
            if show_dialogs:
                messagebox.showerror(self.tr("Error"), message)
            else:
                self.log(message)
                if shared_settings.get("autostart_on_launch"):
                    shared_settings["autostart_on_launch"] = False
                    save_settings(self.settings)
                auto_error = self.tr("SHARED_PRINTER_AUTOSTART_FAILED").format(error=message)
                self.log(auto_error)
            return False

        port_value = self.shared_port_var.get().strip()
        try:
            port_int = int(port_value)
            if not (1 <= port_int <= 65535):
                raise ValueError
        except ValueError:
            message = self.tr("Please enter a valid port number.")
            if show_dialogs:
                messagebox.showerror(self.tr("Error"), message)
            else:
                self.log(message)
                if shared_settings.get("autostart_on_launch"):
                    shared_settings["autostart_on_launch"] = False
                    save_settings(self.settings)
                auto_error = self.tr("SHARED_PRINTER_AUTOSTART_FAILED").format(error=message)
                self.log(auto_error)
            return False

        try:
            self.shared_printer_server.start(port_int, token)
        except Exception as exc:
            if show_dialogs:
                error_message = self.tr("SHARED_PRINTER_START_FAILED").format(error=exc)
                self.log(error_message)
                messagebox.showerror(self.tr("Error"), error_message)
            else:
                auto_error = self.tr("SHARED_PRINTER_AUTOSTART_FAILED").format(error=exc)
                self.log(auto_error)
                if shared_settings.get("autostart_on_launch"):
                    shared_settings["autostart_on_launch"] = False
                    save_settings(self.settings)
            self._set_shared_status("stopped")
            return False

        shared_settings["token"] = token
        shared_settings["port"] = port_int
        shared_settings["autostart_on_launch"] = True
        save_settings(self.settings)

        host = self.shared_printer_server.current_host()
        self._set_shared_status("running", host, port_int)

        if show_dialogs:
            success_message = self.tr("SHARED_PRINTER_STARTED").format(host=host, port=port_int)
            self.log(success_message)
            messagebox.showinfo(self.tr("Information"), success_message)
        else:
            auto_message = self.tr("SHARED_PRINTER_AUTOSTARTED").format(host=host, port=port_int)
            self.log(auto_message)

        return True

    def stop_shared_printer(self) -> None:
        """Arka planda çalışan Flask sunucusunu durdurur."""
        shared_settings = self.settings.setdefault("shared_label_printer", {})
        if shared_settings.get("autostart_on_launch"):
            shared_settings["autostart_on_launch"] = False
            save_settings(self.settings)

        if not self.shared_printer_server.is_running():
            self._set_shared_status("stopped")
            return
        try:
            self.shared_printer_server.stop()
        except Exception as exc:
            self.log(f"{self.tr('Error')}: {exc}")
        self._set_shared_status("stopped")
        stop_message = self.tr("SHARED_PRINTER_STOPPED")
        self.log(stop_message)
        messagebox.showinfo(self.tr("Information"), stop_message)

    def _maybe_auto_start_shared_printer(self) -> None:
        """Uygulama açıldığında gerekirse paylaşılan yazıcıyı başlatır."""
        shared_settings = self.settings.get("shared_label_printer", {})
        if not shared_settings.get("autostart_on_launch"):
            return

        if self.shared_printer_server.is_running():
            return

        self._start_shared_printer(show_dialogs=False)

    def check_shared_printer_status(self) -> None:
        """Sunucunun sağlık durumunu HTTP üzerinden sorgular."""
        token = self.shared_token_var.get().strip()
        if not token:
            messagebox.showerror(self.tr("Error"), self.tr("SHARED_PRINTER_TOKEN_REQUIRED"))
            return

        port_value = self.shared_port_var.get().strip()
        try:
            port_int = int(port_value)
            if not (1 <= port_int <= 65535):
                raise ValueError
        except ValueError:
            messagebox.showerror(self.tr("Error"), self.tr("Please enter a valid port number."))
            return

        self._set_shared_status("checking")
        self.run_in_thread(self._fetch_shared_printer_status, port_int, token)

    def _fetch_shared_printer_status(self, port: int, token: str) -> None:
        """Durum isteğini arka planda gerçekleştirir."""
        url = f"http://127.0.0.1:{port}/status"
        headers = {"Authorization": f"Bearer {token}"}
        try:
            response = requests.get(url, headers=headers, timeout=5)
            if response.status_code == 200:
                data = response.json()
                host = data.get("host") or self.shared_printer_server.current_host()
                remote_port = data.get("port") or port
                self.after(0, lambda: self._set_shared_status("running", host, remote_port))
            else:
                try:
                    payload = response.json()
                    error_text = payload.get("error") or response.text
                except ValueError:
                    error_text = response.text
                self.after(0, lambda: self._handle_status_failure(error_text))
        except Exception as exc:
            self.after(0, lambda: self._handle_status_failure(str(exc)))

    def _handle_status_failure(self, error_message: str) -> None:
        """Durum sorgusu başarısız olduğunda kullanıcıyı bilgilendirir."""
        self._set_shared_status("stopped")
        message = self.tr("SHARED_PRINTER_STATUS_FAILED").format(error=error_message)
        self.log(message)
        messagebox.showerror(self.tr("Error"), message)

    def create_shared_printer_panel(self, parent: ttk.Frame):
        """Paylaşılan yazıcı sekmesini oluşturur."""
        parent.columnconfigure(0, weight=1)
        card = self.create_section_card(parent, "Shared Label Printer")
        card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self._mark_advanced_card(card)
        frame = card.body
        frame.columnconfigure(1, weight=1)

        description = ttk.Label(
            frame,
            text=self.tr("SHARED_PRINTER_DESCRIPTION"),
            style="Description.TLabel",
            wraplength=620,
            justify="left",
        )
        description.grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 12))
        self.register_widget(description, "SHARED_PRINTER_DESCRIPTION")

        token_label = ttk.Label(frame, text=self.tr("Authorization Token:"))
        token_label.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))
        self.register_widget(token_label, "Authorization Token:")

        token_entry = ttk.Entry(frame, textvariable=self.shared_token_var)
        token_entry.grid(row=1, column=1, sticky="we", padx=6, pady=(0, 6))

        port_label = ttk.Label(frame, text=self.tr("Listen Port:"))
        port_label.grid(row=2, column=0, sticky="w", padx=6, pady=(0, 6))
        self.register_widget(port_label, "Listen Port:")

        port_entry = ttk.Entry(frame, textvariable=self.shared_port_var)
        port_entry.grid(row=2, column=1, sticky="we", padx=6, pady=(0, 6))

        status_label = ttk.Label(frame, textvariable=self.shared_status_var)
        status_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 6))

        button_frame = ttk.Frame(frame, style="PanelBody.TFrame")
        button_frame.grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 6))

        start_button = ttk.Button(button_frame, text=self.tr("Start Sharing"), command=self.start_shared_printer)
        start_button.pack(side="left")
        self.register_widget(start_button, "Start Sharing")

        stop_button = ttk.Button(button_frame, text=self.tr("Stop Sharing"), command=self.stop_shared_printer)
        stop_button.pack(side="left", padx=(8, 0))
        self.register_widget(stop_button, "Stop Sharing")

        check_button = ttk.Button(button_frame, text=self.tr("Check Status"), command=self.check_shared_printer_status)
        check_button.pack(side="left", padx=(8, 0))
        self.register_widget(check_button, "Check Status")

        help_frame = ttk.Frame(frame, style="PanelBody.TFrame")
        help_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=6, pady=(12, 6))
        help_frame.grid_columnconfigure(0, weight=1)

        self.shared_help_text = tk.Text(
            help_frame,
            height=5,
            wrap=tk.WORD,
            state=tk.DISABLED,
            background=self.theme_colors["card_bg"],
            foreground=self.theme_colors["text_primary"],
            insertbackground=self.theme_colors["text_primary"],
            relief="flat",
            borderwidth=0,
        )
        self.shared_help_text.grid(row=0, column=0, sticky="nsew")
        try:
            self.shared_help_text.configure(
                disabledforeground=self.theme_colors["text_primary"],
                disabledbackground=self.theme_colors["card_bg"],
            )
        except tk.TclError:
            pass
        scrollbar = ttk.Scrollbar(help_frame, orient="vertical", command=self.shared_help_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.shared_help_text.configure(yscrollcommand=scrollbar.set)

        self._set_shared_status("stopped")

    def on_close(self):
        """Uygulama kapanırken paylaşılan yazıcı sunucusunu durdurur."""
        try:
            try:
                self._persist_view_preferences()
            except Exception:
                pass
            if self.shared_printer_server.is_running():
                try:
                    self.shared_printer_server.stop()
                except Exception as exc:
                    self.log(f"{self.tr('Error')}: {exc}")
        finally:
            self.destroy()

    def create_data_calc_panels(self, parent: ttk.Frame):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        self.format_file = tk.StringVar()
        format_card = self.create_section_card(parent, "4. Format Numbers from File")
        format_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        format_frame = format_card.body
        format_frame.columnconfigure(1, weight=1)

        format_label = ttk.Label(format_frame, text=self.tr("Excel/CSV/TXT File:"))
        format_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(format_label, "Excel/CSV/TXT File:")
        ttk.Entry(format_frame, textvariable=self.format_file).grid(row=0, column=1, sticky="we", padx=6, pady=6)
        format_browse = ttk.Button(
            format_frame,
            text=self.tr("Browse..."),
            command=lambda: self.format_file.set(filedialog.askopenfilename()),
        )
        format_browse.grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.register_widget(format_browse, "Browse...")
        format_button = ttk.Button(format_frame, text=self.tr("Format"), command=self.start_format_numbers)
        format_button.grid(row=0, column=3, padx=6, pady=6)
        self.register_widget(format_button, "Format")

        self.rug_dim_input = tk.StringVar()
        self.rug_result_label = tk.StringVar()
        single_rug_card = self.create_section_card(parent, "5. Rug Size Calculator (Single)")
        single_rug_card.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        single_rug_frame = single_rug_card.body
        single_rug_frame.columnconfigure(1, weight=1)

        rug_label = ttk.Label(single_rug_frame, text=self.tr("Dimension (e.g., 5'2\" x 8'):"))
        rug_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(rug_label, "Dimension (e.g., 5'2\" x 8'):")
        ttk.Entry(single_rug_frame, textvariable=self.rug_dim_input).grid(row=0, column=1, sticky="we", padx=6, pady=6)
        rug_button = ttk.Button(single_rug_frame, text=self.tr("Calculate"), command=self.calculate_single_rug)
        rug_button.grid(row=0, column=2, padx=6, pady=6)
        self.register_widget(rug_button, "Calculate")
        ttk.Label(single_rug_frame, textvariable=self.rug_result_label, font=("Helvetica", 10, "bold")).grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="w",
            padx=6,
            pady=(4, 0),
        )

        self.bulk_rug_file = tk.StringVar()
        self.bulk_rug_col = tk.StringVar(value="Size")
        bulk_rug_card = self.create_section_card(parent, "6. BULK Process Rug Sizes from File")
        bulk_rug_card.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        bulk_rug_frame = bulk_rug_card.body
        bulk_rug_frame.columnconfigure(1, weight=1)

        bulk_file_label = ttk.Label(bulk_rug_frame, text=self.tr("Excel/CSV File:"))
        bulk_file_label.grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.register_widget(bulk_file_label, "Excel/CSV File:")
        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_file).grid(row=0, column=1, padx=6, pady=6, sticky="we")
        bulk_browse = ttk.Button(
            bulk_rug_frame,
            text=self.tr("Browse..."),
            command=lambda: self.bulk_rug_file.set(filedialog.askopenfilename()),
        )
        bulk_browse.grid(row=0, column=2, padx=6, pady=6)
        self.register_widget(bulk_browse, "Browse...")

        bulk_col_label = ttk.Label(bulk_rug_frame, text=self.tr("Column Name/Letter:"))
        bulk_col_label.grid(row=1, column=0, padx=6, pady=6, sticky="w")
        self.register_widget(bulk_col_label, "Column Name/Letter:")
        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_col, width=20).grid(row=1, column=1, padx=6, pady=6, sticky="w")
        bulk_process = ttk.Button(bulk_rug_frame, text=self.tr("Process File"), command=self.start_bulk_rug_sizer)
        bulk_process.grid(row=1, column=2, padx=6, pady=6)
        self.register_widget(bulk_process, "Process File")

        unit_card = self.create_section_card(parent, "7. Unit Converter")
        unit_card.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        unit_frame = unit_card.body
        unit_frame.columnconfigure(1, weight=1)

        self.unit_input = tk.StringVar(value=self.tr("182 cm to ft"))
        self.unit_result_label = tk.StringVar()

        conversion_label = ttk.Label(unit_frame, text=self.tr("Conversion:"))
        conversion_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(conversion_label, "Conversion:")
        ttk.Entry(unit_frame, textvariable=self.unit_input, width=20).grid(row=0, column=1, padx=6, pady=6, sticky="we")
        convert_button = ttk.Button(unit_frame, text=self.tr("Convert"), command=self.convert_units)
        convert_button.grid(row=0, column=2, padx=6, pady=6)
        self.register_widget(convert_button, "Convert")
        ttk.Label(unit_frame, textvariable=self.unit_result_label, font=("Helvetica", 10, "bold")).grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="w",
            padx=6,
            pady=(4, 0),
        )

        wayfair_card = self.create_section_card(parent, "Wayfair Export Formatter")
        wayfair_card.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        wayfair_frame = wayfair_card.body
        wayfair_frame.columnconfigure(0, weight=1)

        self.wayfair_formatter = WayfairFormatter(wayfair_frame, translator=self.tr)
        self.wayfair_formatter.pack(fill="both", expand=True, padx=6, pady=6)

        rug_check_card = self.create_section_card(parent, "Rug No Checker")
        rug_check_card.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        rug_check_frame = rug_check_card.body
        rug_check_frame.columnconfigure(1, weight=1)
        rug_check_frame.rowconfigure(1, weight=1)

        self.rug_mode = tk.StringVar(value="batch")
        self.rug_sold_file = tk.StringVar()
        self.rug_master_file = tk.StringVar()
        self.rug_manual_input = tk.StringVar()
        self.rug_manual_result_var = tk.StringVar()
        self.rug_manual_history_var = tk.StringVar()
        self.rug_comparison_summary_var = tk.StringVar()

        mode_label = ttk.Label(rug_check_frame, text=self.tr("Mode:"))
        mode_label.grid(row=0, column=0, sticky="w", padx=6, pady=(0, 6))
        self.register_widget(mode_label, "Mode:")

        mode_buttons = ttk.Frame(rug_check_frame, style="PanelBody.TFrame")
        mode_buttons.grid(row=0, column=1, sticky="w", padx=6, pady=(0, 6))

        batch_radio = ttk.Radiobutton(
            mode_buttons,
            text=self.tr("Batch Comparison"),
            variable=self.rug_mode,
            value="batch",
            command=self._update_rug_mode,
        )
        batch_radio.pack(side="left", padx=(0, 10))
        self.register_widget(batch_radio, "Batch Comparison")

        manual_radio = ttk.Radiobutton(
            mode_buttons,
            text=self.tr("Manual Search"),
            variable=self.rug_mode,
            value="manual",
            command=self._update_rug_mode,
        )
        manual_radio.pack(side="left")
        self.register_widget(manual_radio, "Manual Search")

        self.rug_batch_frame = ttk.Frame(rug_check_frame, style="PanelBody.TFrame")
        self.rug_batch_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.rug_batch_frame.columnconfigure(1, weight=1)
        self.rug_batch_frame.rowconfigure(4, weight=1)

        sold_label = ttk.Label(self.rug_batch_frame, text=self.tr("Sold List File:"))
        sold_label.grid(row=0, column=0, padx=6, pady=4, sticky="w")
        self.register_widget(sold_label, "Sold List File:")
        ttk.Entry(self.rug_batch_frame, textvariable=self.rug_sold_file, width=50).grid(
            row=0,
            column=1,
            padx=6,
            pady=4,
            sticky="we",
        )
        sold_browse = ttk.Button(
            self.rug_batch_frame,
            text=self.tr("Browse..."),
            command=lambda: self._browse_rug_file(self.rug_sold_file),
        )
        sold_browse.grid(row=0, column=2, padx=6, pady=4)
        self.register_widget(sold_browse, "Browse...")

        master_label = ttk.Label(self.rug_batch_frame, text=self.tr("Master List File:"))
        master_label.grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.register_widget(master_label, "Master List File:")
        ttk.Entry(self.rug_batch_frame, textvariable=self.rug_master_file, width=50).grid(
            row=1,
            column=1,
            padx=6,
            pady=4,
            sticky="we",
        )
        master_browse = ttk.Button(
            self.rug_batch_frame,
            text=self.tr("Browse..."),
            command=lambda: self._browse_rug_file(self.rug_master_file),
        )
        master_browse.grid(row=1, column=2, padx=6, pady=4)
        self.register_widget(master_browse, "Browse...")

        compare_button = ttk.Button(
            self.rug_batch_frame,
            text=self.tr("Start Comparison"),
            command=self.start_rug_comparison,
        )
        compare_button.grid(row=2, column=0, columnspan=3, padx=6, pady=(6, 6), sticky="w")
        self.register_widget(compare_button, "Start Comparison")

        results_label = ttk.Label(self.rug_batch_frame, text=self.tr("Comparison Results:"))
        results_label.grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=(10, 4))
        self.register_widget(results_label, "Comparison Results:")

        results_container = ttk.Frame(self.rug_batch_frame, style="PanelBody.TFrame")
        results_container.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=0, pady=(0, 6))
        results_container.columnconfigure(0, weight=1)
        results_container.rowconfigure(0, weight=1)

        self.rug_result_tree = ttk.Treeview(
            results_container,
            columns=("status", "rug_no"),
            show="headings",
            height=8,
        )
        self.rug_result_tree.heading("status", text=self.tr("Status"))
        self.rug_result_tree.heading("rug_no", text=self.tr("Rug No"))
        self.rug_result_tree.column("status", width=120, anchor="center")
        self.rug_result_tree.column("rug_no", width=200, anchor="w")
        self.rug_result_tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(results_container, orient="vertical", command=self.rug_result_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.rug_result_tree.configure(yscrollcommand=tree_scroll.set)

        summary_label = ttk.Label(self.rug_batch_frame, textvariable=self.rug_comparison_summary_var, style="Secondary.TLabel")
        summary_label.grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))

        self.rug_manual_frame = ttk.Frame(rug_check_frame, style="PanelBody.TFrame")
        self.rug_manual_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.rug_manual_frame.columnconfigure(1, weight=1)

        manual_master_label = ttk.Label(self.rug_manual_frame, text=self.tr("Master List File:"))
        manual_master_label.grid(row=0, column=0, padx=6, pady=4, sticky="w")
        self.register_widget(manual_master_label, "Master List File:")
        ttk.Entry(self.rug_manual_frame, textvariable=self.rug_master_file, width=50).grid(
            row=0,
            column=1,
            padx=6,
            pady=4,
            sticky="we",
        )
        manual_master_browse = ttk.Button(
            self.rug_manual_frame,
            text=self.tr("Browse..."),
            command=lambda: self._browse_rug_file(self.rug_master_file),
        )
        manual_master_browse.grid(row=0, column=2, padx=6, pady=4)
        self.register_widget(manual_master_browse, "Browse...")

        manual_query_label = ttk.Label(self.rug_manual_frame, text=self.tr("Enter Rug No:"))
        manual_query_label.grid(row=1, column=0, padx=6, pady=(10, 4), sticky="w")
        self.register_widget(manual_query_label, "Enter Rug No:")
        ttk.Entry(self.rug_manual_frame, textvariable=self.rug_manual_input, width=30).grid(
            row=1,
            column=1,
            padx=6,
            pady=(10, 4),
            sticky="we",
        )
        manual_search_button = ttk.Button(
            self.rug_manual_frame,
            text=self.tr("Search"),
            command=self.start_manual_rug_search,
        )
        manual_search_button.grid(row=1, column=2, padx=6, pady=(10, 4))
        self.register_widget(manual_search_button, "Search")

        manual_result_label = ttk.Label(self.rug_manual_frame, textvariable=self.rug_manual_result_var, style="Secondary.TLabel")
        manual_result_label.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(6, 4))

        manual_history_title = ttk.Label(self.rug_manual_frame, text=self.tr("Manual Search History:"))
        manual_history_title.grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=(10, 2))
        self.register_widget(manual_history_title, "Manual Search History:")

        manual_history_label = ttk.Label(
            self.rug_manual_frame,
            textvariable=self.rug_manual_history_var,
            justify="left",
            style="Secondary.TLabel",
        )
        manual_history_label.grid(row=4, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))

        self.rug_manual_frame.grid_remove()
        self._update_manual_history_display()
        self._update_manual_result_label()
        self._update_rug_mode()
        self._refresh_rug_comparison_display()

        image_link_card = self.create_section_card(parent, "8. Match Image Links")
        image_link_card.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        self._mark_advanced_card(image_link_card)
        image_link_frame = image_link_card.body
        image_link_frame.columnconfigure(1, weight=1)

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

    def create_rug_no_control_tab(self, parent: ttk.Frame) -> ttk.Frame:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        self.rug_control_sold_path = tk.StringVar()
        self.rug_control_inventory_path = tk.StringVar()

        input_frame = ttk.Frame(parent, style="PanelBody.TFrame")
        input_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(4, weight=1)

        sold_label = ttk.Label(input_frame, text=self.tr("Sold List File:"))
        sold_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(sold_label, "Sold List File:")
        ttk.Entry(input_frame, textvariable=self.rug_control_sold_path).grid(
            row=0,
            column=1,
            columnspan=3,
            sticky="we",
            padx=6,
            pady=6,
        )
        sold_browse = ttk.Button(
            input_frame,
            text=self.tr("Browse..."),
            command=lambda: self._browse_rug_file(self.rug_control_sold_path),
        )
        sold_browse.grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self.register_widget(sold_browse, "Browse...")

        inventory_label = ttk.Label(input_frame, text=self.tr("Inventory List File:"))
        inventory_label.grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.register_widget(inventory_label, "Inventory List File:")
        ttk.Entry(input_frame, textvariable=self.rug_control_inventory_path).grid(
            row=1,
            column=1,
            columnspan=3,
            sticky="we",
            padx=6,
            pady=6,
        )
        inventory_browse = ttk.Button(
            input_frame,
            text=self.tr("Browse..."),
            command=lambda: self._browse_rug_file(self.rug_control_inventory_path),
        )
        inventory_browse.grid(row=1, column=4, sticky="e", padx=6, pady=6)
        self.register_widget(inventory_browse, "Browse...")

        check_button = ttk.Button(
            input_frame,
            text=self.tr("Check Rug Nos"),
            command=self.run_rug_no_control_check,
        )
        check_button.grid(row=2, column=0, columnspan=5, sticky="w", padx=6, pady=(6, 0))
        self.register_widget(check_button, "Check Rug Nos")

        results_label = ttk.Label(parent, text=self.tr("Results:"), style="Secondary.TLabel")
        results_label.grid(row=1, column=0, sticky="w", padx=12)
        self.register_widget(results_label, "Results:")
        self.rug_control_results_label = results_label

        tree_container = ttk.Frame(parent, style="PanelBody.TFrame")
        tree_container.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        columns = ("rug_no", "status")
        tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=12)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        self.rug_control_tree = tree
        self.populate_rug_no_control_tree(self.rug_control_results)

        return parent

    def run_rug_no_control_check(self) -> None:
        sold_path = self.rug_control_sold_path.get().strip()
        inventory_path = self.rug_control_inventory_path.get().strip()
        if not sold_path or not inventory_path:
            messagebox.showerror(
                self.tr("Error"),
                self.tr("Please select both Sold and Inventory files."),
            )
            return

        try:
            results = self.load_rug_no_control_data(sold_path, inventory_path)
        except FileNotFoundError as exc:
            missing = getattr(exc, "filename", str(exc)) or str(exc)
            message = self.tr("Could not read the selected file: {error}").format(error=missing)
            messagebox.showerror(self.tr("Error"), message)
            return
        except ValueError as exc:
            messagebox.showerror(self.tr("Error"), str(exc))
            return

        self.rug_control_results = results
        self.populate_rug_no_control_tree(results)
        self.log(self.tr("Rug No control completed."))

    def load_rug_no_control_data(self, sold_path: str, inventory_path: str) -> List[Tuple[str, bool]]:
        sold_values = self._load_sold_rug_numbers(sold_path)
        inventory_values = self._load_inventory_rug_numbers(inventory_path)
        return [(original, normalized in inventory_values) for original, normalized in sold_values]

    def _read_rug_no_control_dataframe(self, path: str) -> pd.DataFrame:
        extension = os.path.splitext(path)[1].lower()
        try:
            if extension in {".xlsx", ".xls", ".xlsm", ".xlsb"}:
                dataframe = pd.read_excel(path, dtype=str)
            else:
                dataframe = pd.read_csv(path, dtype=str, keep_default_na=False)
        except FileNotFoundError:
            raise
        except Exception as exc:  # pylint: disable=broad-except
            message = self.tr("Could not read the selected file: {error}").format(error=exc)
            raise ValueError(message) from exc

        if not isinstance(dataframe, pd.DataFrame):
            dataframe = pd.DataFrame(dataframe)

        return dataframe.fillna("")

    def _find_rug_no_columns(self, dataframe: pd.DataFrame) -> List[str]:
        normalized_candidates = {candidate.strip().lower() for candidate in RUG_NO_CONTROL_COLUMNS}
        matches = []
        for column in dataframe.columns:
            normalized = str(column).strip().lower()
            if normalized in normalized_candidates:
                matches.append(column)
        return matches

    def _extract_rug_values(self, series: pd.Series) -> List[Tuple[str, str]]:
        values: List[Tuple[str, str]] = []
        for raw in series:
            if pd.isna(raw):
                continue
            text = str(raw).strip()
            if not text:
                continue
            values.append((text, text.lower()))
        return values

    def _load_sold_rug_numbers(self, path: str) -> List[Tuple[str, str]]:
        dataframe = self._read_rug_no_control_dataframe(path)
        matches = self._find_rug_no_columns(dataframe)
        if not matches:
            raise ValueError(self.tr("Could not find a Rug No column in the selected file."))

        primary_column = matches[0]
        return self._extract_rug_values(dataframe[primary_column])

    def _load_inventory_rug_numbers(self, path: str) -> Set[str]:
        dataframe = self._read_rug_no_control_dataframe(path)
        matches = self._find_rug_no_columns(dataframe)
        if not matches:
            raise ValueError(self.tr("Could not find a Rug No column in the selected file."))

        normalized_values: Set[str] = set()
        for column in matches:
            for _original, normalized in self._extract_rug_values(dataframe[column]):
                normalized_values.add(normalized)
        return normalized_values

    def populate_rug_no_control_tree(self, results: List[Tuple[str, bool]]):
        tree = getattr(self, "rug_control_tree", None)
        if tree is None:
            return None

        tree.heading("rug_no", text=self.tr("Rug No"))
        tree.heading("status", text=self.tr("Status"))

        for item in tree.get_children():
            tree.delete(item)

        for original, found in results:
            status_text = self.tr("RUG_NO_CONTROL_FOUND") if found else self.tr("RUG_NO_CONTROL_NOT_FOUND")
            tree.insert("", "end", values=(original, status_text))

        return tree

    def create_code_gen_panels(self, parent: ttk.Frame):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        def toggle_dymo_options(output_var, combobox, entry):
            if output_var.get() == "Dymo":
                combobox.config(state="readonly")
                entry.config(state="normal")
            else:
                combobox.config(state="disabled")
                entry.config(state="disabled")

        qr_card = self.create_section_card(parent, "8. QR Code Generator")
        qr_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self._mark_advanced_card(qr_card)
        qr_frame = qr_card.body
        qr_frame.columnconfigure(1, weight=1)
        qr_frame.columnconfigure(3, weight=1)

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

        qr_radio_frame = ttk.Frame(qr_frame, style="PanelBody.TFrame")
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

        bc_card = self.create_section_card(parent, "9. Barcode Generator")
        bc_card.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self._mark_advanced_card(bc_card)
        bc_frame = bc_card.body
        bc_frame.columnconfigure(1, weight=1)
        bc_frame.columnconfigure(3, weight=1)

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

        bc_radio_frame = ttk.Frame(bc_frame, style="PanelBody.TFrame")
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

    def create_rinven_tag_panel(self, parent: ttk.Frame):
        parent.columnconfigure(0, weight=1)
        card = self.create_section_card(parent, "Rinven Tag")
        card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self._mark_advanced_card(card)
        frame = card.body
        frame.columnconfigure(1, weight=1)

        self.rinven_collection = tk.StringVar()
        self.rinven_design = tk.StringVar()
        self.rinven_color = tk.StringVar()
        self.rinven_size = tk.StringVar()
        self.rinven_origin = tk.StringVar()
        self.rinven_style = tk.StringVar()
        self.rinven_content = tk.StringVar()
        self.rinven_type = tk.StringVar()
        self.rinven_rug_no = tk.StringVar()
        self.rinven_barcode_data = tk.StringVar()
        self.rinven_filename = tk.StringVar(value="rinven_tag.png")
        self.rinven_include_barcode = tk.BooleanVar(value=False)

        self.rinven_field_widgets = {}

        fields = [
            ("collection", "Collection Name:", self.rinven_collection),
            ("design", "Design:", self.rinven_design),
            ("color", "Color:", self.rinven_color),
            ("size", "Size:", self.rinven_size),
            ("origin", "Origin:", self.rinven_origin),
            ("style", "Style:", self.rinven_style),
            ("content", "Content:", self.rinven_content),
            ("type", "Type:", self.rinven_type),
            ("rug_no", "Rug #:", self.rinven_rug_no),
        ]

        for row, (field_key, label_key, var) in enumerate(fields):
            label = ttk.Label(frame, text=self.tr(label_key))
            label.grid(row=row, column=0, sticky="e", padx=6, pady=4)
            self.register_widget(label, label_key)
            history_values = self.settings.get("rinven_history", {}).get(field_key, [])
            combobox = ttk.Combobox(
                frame,
                textvariable=var,
                values=history_values,
                style="Light.TCombobox",
            )
            combobox.grid(row=row, column=1, sticky="we", padx=6, pady=4)
            self.rinven_field_widgets[field_key] = combobox

        row_offset = len(fields)

        barcode_check = ttk.Checkbutton(
            frame,
            text=self.tr("Include Barcode"),
            variable=self.rinven_include_barcode,
            command=self.toggle_rinven_barcode,
        )
        barcode_check.grid(row=row_offset, column=0, columnspan=2, sticky="w", padx=6, pady=(12, 4))
        self.register_widget(barcode_check, "Include Barcode")

        barcode_label = ttk.Label(frame, text=self.tr("Barcode Data:"))
        barcode_label.grid(row=row_offset + 1, column=0, sticky="e", padx=6, pady=4)
        self.register_widget(barcode_label, "Barcode Data:")

        self.rinven_barcode_entry = ttk.Entry(frame, textvariable=self.rinven_barcode_data, state="disabled")
        self.rinven_barcode_entry.grid(row=row_offset + 1, column=1, sticky="we", padx=6, pady=4)

        filename_label = ttk.Label(frame, text=self.tr("Filename:"))
        filename_label.grid(row=row_offset + 2, column=0, sticky="e", padx=6, pady=4)
        self.register_widget(filename_label, "Filename:")

        ttk.Entry(frame, textvariable=self.rinven_filename).grid(
            row=row_offset + 2, column=1, sticky="we", padx=6, pady=4
        )

        generate_button = ttk.Button(
            frame,
            text=self.tr("Generate Rinven Tag"),
            command=self.start_generate_rinven_tag,
        )
        generate_button.grid(row=row_offset + 3, column=0, columnspan=2, pady=(12, 6))
        self.register_widget(generate_button, "Generate Rinven Tag")

    def create_about_panel(self, parent: ttk.Frame):
        parent.columnconfigure(0, weight=1)
        card = self.create_section_card(parent, "Help & About")
        card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        frame = card.body

        top_frame = ttk.Frame(frame, style="PanelBody.TFrame")
        top_frame.pack(fill="x", padx=0, pady=5)

        update_button = ttk.Button(top_frame, text=self.tr("Check for Updates"), command=lambda: self.run_in_thread(check_for_updates, self, self.log, __version__, silent=False))
        update_button.pack(side="left")
        self.register_widget(update_button, "Check for Updates")

        self.help_text_area = ScrolledText(
            frame,
            wrap=tk.WORD,
            padx=10,
            pady=10,
            font=("Helvetica", self._scaled_size(10)),
        )
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

    def _browse_rug_file(self, variable: tk.StringVar) -> None:
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("All Files", "*.*"),
            ]
        )
        if file_path:
            variable.set(file_path)

    def _update_rug_mode(self) -> None:
        if not hasattr(self, "rug_mode"):
            return
        mode = self.rug_mode.get()
        if mode == "manual":
            self.rug_batch_frame.grid_remove()
            self.rug_manual_frame.grid()
        else:
            self.rug_manual_frame.grid_remove()
            self.rug_batch_frame.grid()

    def start_rug_comparison(self) -> None:
        sold_path = self.rug_sold_file.get().strip()
        master_path = self.rug_master_file.get().strip()
        if not sold_path or not master_path:
            messagebox.showerror(
                self.tr("Error"),
                self.tr("Please select both Sold List and Master List files."),
            )
            return

        self.rug_comparison_results = None
        self._refresh_rug_comparison_display()

        self.run_in_thread(
            backend.compare_rug_numbers_task,
            sold_path,
            master_path,
            self.log,
            self._handle_rug_comparison_completion,
            self.handle_rug_comparison_results,
        )

    def _handle_rug_comparison_completion(self, status: str, message: str) -> None:
        translated_message = self.tr(message) if message else message
        self.task_completion_popup(status, translated_message)

    def handle_rug_comparison_results(self, found, missing) -> None:
        def update():
            self.rug_comparison_results = (list(found), list(missing))
            self._refresh_rug_comparison_display()

        self.after(0, update)

    def _refresh_rug_comparison_display(self) -> None:
        tree = getattr(self, "rug_result_tree", None)
        if tree is None:
            return

        tree.heading("status", text=self.tr("Status"))
        tree.heading("rug_no", text=self.tr("Rug No"))

        for item in tree.get_children():
            tree.delete(item)

        results = getattr(self, "rug_comparison_results", None)
        if results is None:
            self.rug_comparison_summary_var.set("")
            return

        found, missing = results
        status_found = self.tr("FOUND")
        status_missing = self.tr("MISSING")

        for number in found:
            tree.insert("", "end", values=(status_found, number))
        for number in missing:
            tree.insert("", "end", values=(status_missing, number))

        summary = self.tr("Found: {found} | Missing: {missing}").format(
            found=len(found), missing=len(missing)
        )
        self.rug_comparison_summary_var.set(summary)

    def start_manual_rug_search(self) -> None:
        master_path = self.rug_master_file.get().strip()
        if not master_path:
            messagebox.showerror(
                self.tr("Error"),
                self.tr("Please select a Master List file."),
            )
            return

        query_raw = self.rug_manual_input.get()
        normalized_query = backend.normalize_rug_number(query_raw)
        if not normalized_query:
            messagebox.showerror(self.tr("Error"), self.tr("Please enter a Rug No."))
            return

        self.rug_manual_input.set(normalized_query)
        self.run_in_thread(self._manual_rug_search_worker, master_path, normalized_query)

    def _manual_rug_search_worker(self, master_path: str, rug_no: str) -> None:
        try:
            numbers = backend.load_rug_numbers_from_file(master_path)
        except Exception as exc:
            message = str(exc)
            self.log(self.tr("Error: {message}").format(message=message))

            def show_error() -> None:
                messagebox.showerror(self.tr("Error"), message)

            self.after(0, show_error)
            return

        found = rug_no in set(numbers)

        self.after(0, lambda: self._apply_manual_search_result(rug_no, found))

    def _apply_manual_search_result(self, rug_no: str, found: bool) -> None:
        # Update last result and keep only the latest five history entries.
        self.rug_manual_last_result = (rug_no, found)

        history = [entry for entry in self.rug_manual_history_entries if entry[0] != rug_no]
        history.insert(0, (rug_no, found))
        if len(history) > 5:
            history = history[:5]
        self.rug_manual_history_entries = history

        self._update_manual_result_label()
        self._update_manual_history_display()

    def _update_manual_result_label(self) -> None:
        if getattr(self, "rug_manual_result_var", None) is None:
            return

        result = getattr(self, "rug_manual_last_result", None)
        if not result:
            self.rug_manual_result_var.set("")
            return

        rug_no, found = result
        template = (
            "Rug No {number} found in master list."
            if found
            else "Rug No {number} not found in master list."
        )
        self.rug_manual_result_var.set(self.tr(template).format(number=rug_no))

    def _update_manual_history_display(self) -> None:
        if getattr(self, "rug_manual_history_var", None) is None:
            return

        if self.rug_manual_history_entries:
            lines = []
            for rug_no, found in self.rug_manual_history_entries:
                status_text = self.tr("Found") if found else self.tr("Not Found")
                lines.append(f"{rug_no} — {status_text}")
            self.rug_manual_history_var.set("\n".join(lines))
        else:
            self.rug_manual_history_var.set(self.tr("No recent searches yet."))

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

    def toggle_rinven_barcode(self):
        if self.rinven_include_barcode.get():
            self.rinven_barcode_entry.config(state="normal")
        else:
            self.rinven_barcode_entry.config(state="disabled")
            self.rinven_barcode_data.set("")

    def update_rinven_history(self, details: dict):
        history = self.settings.setdefault("rinven_history", {})
        updated = False

        for field_key, value in details.items():
            if not value:
                continue

            stored_values = history.setdefault(field_key, [])

            if value in stored_values:
                if stored_values[0] != value:
                    stored_values.remove(value)
                    stored_values.insert(0, value)
                    updated = True
            else:
                stored_values.insert(0, value)
                updated = True

            if len(stored_values) > 10:
                del stored_values[10:]

        if updated:
            save_settings(self.settings)
            for key, combobox in self.rinven_field_widgets.items():
                combobox["values"] = history.get(key, [])

    def start_generate_rinven_tag(self):
        details = {
            "collection": self.rinven_collection.get().strip(),
            "design": self.rinven_design.get().strip(),
            "color": self.rinven_color.get().strip(),
            "size": self.rinven_size.get().strip(),
            "origin": self.rinven_origin.get().strip(),
            "style": self.rinven_style.get().strip(),
            "content": self.rinven_content.get().strip(),
            "type": self.rinven_type.get().strip(),
            "rug_no": self.rinven_rug_no.get().strip(),
        }

        if not all(details.values()):
            messagebox.showerror(self.tr("Error"), self.tr("Please fill in all Rinven Tag fields."))
            return

        filename = self.rinven_filename.get().strip()
        if not filename:
            messagebox.showerror(self.tr("Error"), self.tr("Filename is required."))
            return

        include_barcode = self.rinven_include_barcode.get()
        barcode_value = self.rinven_barcode_data.get().strip()

        if include_barcode and not barcode_value:
            messagebox.showerror(self.tr("Error"), self.tr("Barcode data is required when barcode is enabled."))
            return

        self.update_rinven_history(details)

        log_msg, success_msg = backend.generate_rinven_tag_label(
            details,
            filename,
            include_barcode,
            barcode_value,
        )
        self.log(log_msg)
        if success_msg:
            self.task_completion_popup("Success", success_msg)
        else:
            messagebox.showerror(self.tr("Error"), log_msg)


def get_rug_no_control_functions() -> Tuple[Callable, Callable, Callable]:
    """Return helper callables for the Rug No Control tab."""

    return (
        ToolApp.create_rug_no_control_tab,
        ToolApp.load_rug_no_control_data,
        ToolApp.populate_rug_no_control_tree,
    )

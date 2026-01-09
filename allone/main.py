# main.py
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

# Arayüz sınıfımızı app_ui dosyasından import ediyoruz.
from app_ui import ToolApp

def install_and_check():
    """Checks for required libraries and installs them if they are missing."""
    required_packages = [
        'tqdm', 'openpyxl', 'Pillow', 'pillow-heif',
        'pandas', 'requests', 'xlrd', 'qrcode', 'python-barcode',
        'numpy'
    ]
    print("Checking for required libraries...")
    for package in required_packages:
        try:
            import_name = package.replace('-', '_')
            __import__(import_name)
        except ImportError:
            try:
                print(f"'{package}' not found. Installing...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            except subprocess.CalledProcessError:
                print(f"ERROR: Failed to install '{package}'. Please install manually.")
                sys.exit(1)
    print("\n✅ Setup checks complete. Starting GUI...")


if __name__ == "__main__":
    # Eğer program bir .exe olarak derlenmişse, bu kontrolü atla.
    # 'sys.frozen' özelliği sadece PyInstaller .exe'lerinde bulunur.
    if not getattr(sys, 'frozen', False):
        install_and_check()
    
    # Uygulama arayüzünü başlat.
    app = ToolApp()
    app.mainloop()

# setup.py
import os
import sys
import subprocess
import urllib.request

# ANSI renk kodları (daha güzel bir görünüm için)
class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    END = '\033[0m'

def check_and_install_requests():
    """Checks if 'requests' is installed, if not, installs it."""
    try:
        import requests
    except ImportError:
        print(f"{Colors.YELLOW}Python 'requests' library not found. Installing...{Colors.END}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "requests"])
            print(f"{Colors.GREEN}Successfully installed 'requests'.{Colors.END}")
        except subprocess.CalledProcessError:
            print(f"{Colors.RED}ERROR: Failed to install 'requests'. Please install it manually ('pip install requests').{Colors.END}")
            sys.exit(1)

def download_file(url, filename, description):
    """Downloads a file from a URL and saves it locally."""
    try:
        print(f"Downloading {Colors.BOLD}{description}{Colors.END} ({filename})... ", end="", flush=True)
        # urllib.request.urlretrieve(url, filename) # Bu bazen firewall'lara takılabilir
        with urllib.request.urlopen(url) as response, open(filename, 'wb') as out_file:
            data = response.read() # a `bytes` object
            out_file.write(data)
        print(f"{Colors.GREEN}Done.{Colors.END}")
        return True
    except Exception as e:
        print(f"{Colors.RED}Failed.\nError: {e}{Colors.END}")
        return False

def main():
    """Main setup function."""
    print(f"\n{Colors.BOLD}--- Combined Utility Tool Setup ---{Colors.END}")
    
    # 1. Gerekli 'requests' kütüphanesini kontrol et/yükle
    # Gerçi bu betik artık urllib kullanıyor ama yine de iyi bir pratik.
    # check_and_install_requests() # requests'e gerek kalmadı.
    
    # 2. GitHub'dan indirilecek dosyaları tanımla
    github_user = "hakanakaslanx-code"
    repo_name = "allone"
    branch = "main"
    # DİKKAT: Eğer dosyalar bir alt klasördeyse, buraya ekle. Örneğin: "allone/"
    # Şu anki yapıya göre ana dizinde olduğunu varsayıyoruz.
    subfolder = "allone/" 

    files_to_download = {
        "main.py": "Main Application Launcher",
        "app_ui.py": "Application UI",
        "backend_logic.py": "Backend Logic",
        "settings_manager.py": "Settings Manager",
        "updater.py": "Auto-Updater"
    }
    
    all_successful = True
    for filename, description in files_to_download.items():
        url = f"https://raw.githubusercontent.com/{github_user}/{repo_name}/{branch}/{subfolder}{filename}"
        if not download_file(url, filename, description):
            all_successful = False
            break
            
    if not all_successful:
        print(f"\n{Colors.RED}One or more files failed to download. Setup cannot continue.{Colors.END}")
        sys.exit(1)
        
    print(f"\n{Colors.GREEN}All application files have been downloaded successfully.{Colors.END}")
    
    # 3. Gerekli Python kütüphanelerini yüklemek için main.py'yi çalıştır
    print(f"\n{Colors.YELLOW}Running dependency check and starting the application...{Colors.END}")
    print("This may take a moment if libraries need to be installed.")
    print("-" * 40)
    
    try:
        # Bu komut, main.py'yi çalıştırır. 
        # main.py içindeki install_and_check() fonksiyonu kütüphaneleri yükler,
        # ardından app.mainloop() uygulamayı başlatır.
        subprocess.run([sys.executable, "main.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"\n{Colors.RED}An error occurred while running the application: {e}{Colors.END}")
        print("Please try running 'python main.py' manually.")
    except FileNotFoundError:
        print(f"\n{Colors.RED}Error: 'main.py' could not be found after download.{Colors.END}")

if __name__ == "__main__":
    main()
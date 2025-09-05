# Bu dosya, eski uygulamanın yerine geçecek olan özel yükleyicidir.
# Adını GitHub'a yüklerken "allone.beta-en.py" yapacağız.

__version__ = "3.0"  # Versiyonu çok yüksek tutuyoruz ki tüm eski versiyonlar bunu güncelleme olarak algılasın.

import os
import sys
import subprocess
import urllib.request
import time

# ANSI renk kodları
class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'

def download_file(url, filename, description):
    """Downloads a file from a URL and saves it locally."""
    try:
        print(f"  Downloading {Colors.BOLD}{description}{Colors.END} ({filename})... ", end="", flush=True)
        with urllib.request.urlopen(url) as response, open(filename, 'wb') as out_file:
            data = response.read()
            out_file.write(data)
        print(f"{Colors.GREEN}Done.{Colors.END}")
        return True
    except Exception as e:
        print(f"{Colors.RED}Failed.\n  Error: {e}{Colors.END}")
        return False

def main():
    """Main setup function."""
    print(f"\n{Colors.CYAN}{Colors.BOLD}--- Application Upgrade Required ---{Colors.END}")
    print(f"{Colors.YELLOW}Your application is being upgraded to a new, multi-file structure.{Colors.END}")
    print("This one-time process will download the new version.")
    
    # 1. Yeni uygulama için bir alt klasör oluştur
    app_folder = "AllONE"
    if not os.path.exists(app_folder):
        os.makedirs(app_folder)
    
    os.chdir(app_folder)
    print(f"\nCreated a new folder for the application: {Colors.BOLD}{app_folder}{Colors.END}")

    # 2. GitHub'dan indirilecek dosyaları tanımla
    github_user = "hakanakaslanx-code"
    repo_name = "allone"
    branch = "main"
    subfolder = "allone/" 

    files_to_download = {
        "main.py": "Main Application Launcher",
        "app_ui.py": "Application UI",
        "backend_logic.py": "Backend Logic",
        "settings_manager.py": "Settings Manager",
        "updater.py": "Auto-Updater"
    }
    
    print("\nDownloading new application files:")
    all_successful = True
    for filename, description in files_to_download.items():
        url = f"https://raw.githubusercontent.com/{github_user}/{repo_name}/{branch}/{subfolder}{filename}"
        if not download_file(url, filename, description):
            all_successful = False
            break
            
    if not all_successful:
        print(f"\n{Colors.RED}One or more files failed to download. Please try again later.{Colors.END}")
        time.sleep(5)
        sys.exit(1)
        
    print(f"\n{Colors.GREEN}All application files have been downloaded successfully.{Colors.END}")
    
    # 3. Gerekli Python kütüphanelerini yükle ve uygulamayı başlat
    print(f"\n{Colors.YELLOW}Running final setup and starting the new application...{Colors.END}")
    print("-" * 50)
    
    try:
        # main.py'yi çalıştırarak hem kütüphaneleri kur hem de uygulamayı başlat
        subprocess.run([sys.executable, "main.py"])
    except Exception as e:
        print(f"\n{Colors.RED}An error occurred while launching the new application: {e}{Colors.END}")
        print(f"You can start the application manually by running 'python main.py' inside the '{app_folder}' folder.")
        time.sleep(10)

    print(f"\n{Colors.GREEN}{Colors.BOLD}Upgrade complete!{Colors.END}")
    print("You can now delete the old application file. Please use the new application from now on.")
    time.sleep(5)


if __name__ == "__main__":
    main()

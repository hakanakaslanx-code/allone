# downloader.py
import os
import sys
import subprocess
import urllib.request
import time

class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'

def download_file(url, filename, description):
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
    print(f"\n{Colors.CYAN}{Colors.BOLD}--- Application Update In Progress ---{Colors.END}")
    print("Downloading the latest version of all application files...")
    
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
    
    # İndirme işleminin yapılacağı ana dizine geç
    # Bu betik, ana uygulama ile aynı dizinde olmalı
    # Eğer değilse, bu kısmı ayarlamak gerekebilir.
    # Şimdilik aynı dizinde olduğunu varsayıyoruz.

    all_successful = True
    for filename, description in files_to_download.items():
        url = f"https://raw.githubusercontent.com/{github_user}/{repo_name}/{branch}/{subfolder}{filename}"
        if not download_file(url, filename, description):
            all_successful = False
            break
            
    if not all_successful:
        print(f"\n{Colors.RED}Update failed. One or more files could not be downloaded.{Colors.END}")
        time.sleep(5)
        sys.exit(1)
        
    print(f"\n{Colors.GREEN}All application files have been updated successfully.{Colors.END}")
    
    print(f"\n{Colors.YELLOW}Restarting the application...{Colors.END}")
    print("-" * 40)
    
    try:
        # Popen kullanmak, bu betiğin kapanmasına izin verirken uygulamanın çalışmasını sağlar.
        subprocess.Popen([sys.executable, "main.py"])
    except Exception as e:
        print(f"\n{Colors.RED}An error occurred while restarting the application: {e}{Colors.END}")
        print("Please start the application manually by running 'python main.py'.")
        time.sleep(10)

if __name__ == "__main__":
    main()
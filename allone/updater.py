# updater.py
import os
import sys
import re
import requests
import subprocess
from tkinter import messagebox

def check_for_updates(app_instance, log_callback, current_version, silent=False):
    log_callback("Checking for updates...")
    
    # Versiyon kontrolü için app_ui.py'yi kullanıyoruz
    version_check_url = "https://raw.githubusercontent.com/hakanakaslanx-code/allone/main/allone/app_ui.py"
    
    # Güncelleme için ise downloader.py'yi indireceğiz
    downloader_url = "https://raw.githubusercontent.com/hakanakaslanx-code/allone/main/allone/downloader.py"
    
    try:
        response = requests.get(version_check_url, timeout=10)
        response.raise_for_status()
        remote_script_content = response.text
        
        match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", remote_script_content)
        if not match:
            if not silent: messagebox.showwarning("Update Check", "Could not determine remote version.")
            return
        
        remote_version = match.group(1)
        
        if remote_version > current_version:
            log_callback(f"--> New version ({remote_version}) available.")
            if messagebox.askyesno("Update Available", f"A new version ({remote_version}) is available.\n\nThis will download all new files and restart the application.\n\nDo you want to update now?"):
                log_callback("Downloading updater...")
                
                # Yeni downloader.py'yi indir ve geçici bir dosyaya kaydet
                downloader_content = requests.get(downloader_url).text
                downloader_path = "downloader_update.py"
                with open(downloader_path, "w", encoding='utf-8') as f:
                    f.write(downloader_content)
                
                log_callback("Update started. The application will now close.")
                
                # Ana uygulamayı kapat ve downloader'ı çalıştır
                # Downloader daha sonra eski downloader_update.py'yi silebilir.
                subprocess.Popen([sys.executable, downloader_path])
                app_instance.destroy()
                sys.exit(0)
        else:
            log_callback(f"✅ Your application is up-to-date. (Version: {current_version})")
            if not silent: messagebox.showinfo("Update Check", "You are running the latest version.")
            
    except Exception as e:
        log_callback(f"Warning: Update check failed. Reason: {e}")
        if not silent: messagebox.showerror("Update Check Failed", f"Could not check for updates.\n\nReason: {e}")

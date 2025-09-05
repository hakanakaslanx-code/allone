# updater.py
import os
import sys
import re
import requests
import subprocess
from tkinter import messagebox

def check_for_updates(app_instance, log_callback, current_version, silent=False):
    """Checks for a new version of the script on GitHub and self-updates if one is found."""
    log_callback("Checking for updates...")
    # ÖNEMLİ: Bu URL'nin güncel GUI kodunu içeren dosyaya işaret ettiğinden emin olun.
    script_url = "https://raw.githubusercontent.com/hakanakaslanx-code/allone/main/allone/app_ui.py"
    current_script_name = os.path.basename(sys.argv[0])
    
    try:
        response = requests.get(script_url, timeout=10)
        response.raise_for_status()
        remote_script_content = response.text
        
        match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", remote_script_content)
        if not match:
            log_callback("Warning: Could not determine remote version. Skipping update.")
            if not silent: messagebox.showwarning("Update Check", "Could not determine remote version number.")
            return
        
        remote_version = match.group(1)
        
        if remote_version > current_version:
            log_callback(f"--> New version ({remote_version}) available.")
            if messagebox.askyesno("Update Available", f"A new version ({remote_version}) is available.\nYour current version is {current_version}.\n\nDo you want to update now?"):
                messagebox.showinfo("Updating...", "The application will now close, update itself, and restart.")
                log_callback("Updating...")
                new_script_path = current_script_name + ".new"
                with open(new_script_path, "w", encoding='utf-8') as f: f.write(remote_script_content)
                
                if sys.platform == 'win32':
                    updater_path = "updater.bat"
                    content = f'@echo off\ntimeout /t 2 /nobreak > NUL\ndel "{current_script_name}"\nrename "{new_script_path}" "{current_script_name}"\nstart "" "{sys.executable}" "{current_script_name}"\ndel "{updater_path}"'
                else:
                    updater_path = "updater.sh"
                    content = f'#!/bin/bash\nsleep 2\nrm "{current_script_name}"\nmv "{new_script_path}" "{current_script_name}"\nchmod +x "{current_script_name}"\n"{sys.executable}" "{current_script_name}" &\nrm -- "$0"'
                
                with open(updater_path, "w", encoding='utf-8') as f: f.write(content)
                if sys.platform != 'win32': os.chmod(updater_path, 0o755)
                
                subprocess.Popen([updater_path])
                app_instance.destroy()
                sys.exit(0)
        else:
            log_callback(f"✅ Your application is up-to-date. (Version: {current_version})")
            if not silent: messagebox.showinfo("Update Check", "You are running the latest version.")
            
    except Exception as e:
        log_callback(f"Warning: Update check failed. Reason: {e}")

        if not silent: messagebox.showerror("Update Check Failed", f"Could not check for updates.\n\nReason: {e}")

# updater.py
import os
import sys
import re
import requests
import webbrowser
from tkinter import messagebox

def check_for_updates(app_instance, log_callback, current_version, silent=False):
    log_callback("Checking for updates...")
    
    owner = "hakanakaslanx-code"
    repo = "allone"
    
    version_check_url = f"https://raw.githubusercontent.com/{owner}/{repo}/main/allone/app_ui.py"
    
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
            
            latest_release_url = f"https://github.com/{owner}/{repo}/releases/latest"
            
            if messagebox.askyesno("Update Available", f"A new version ({remote_version}) is available.\n\nDo you want to open the download page now?"):
                log_callback(f"Opening download page: {latest_release_url}")
                webbrowser.open(latest_release_url)
                
                messagebox.showinfo("Update", 
                                  "The download page has been opened in your browser.\n\n"
                                  "Please download the new 'AllOneTool.exe' and replace your old application with it.")
                
                app_instance.destroy()
                sys.exit(0)

        else:
            log_callback(f"✅ Your application is up-to-date. (Version: {current_version})")
            if not silent: messagebox.showinfo("Update Check", "You are running the latest version.")
            
    except Exception as e:
        log_callback(f"Warning: Update check failed. Reason: {e}")
        if not silent: messagebox.showerror("Update Check Failed", f"Could not check for updates.\n\nReason: {e}")

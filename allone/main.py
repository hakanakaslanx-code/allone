# main.py
import subprocess
import sys

def install_and_check():
    """No-op. Dependencies are bundled in frozen mode."""
    pass


if __name__ == "__main__":
    # Eğer program bir .exe olarak derlenmişse, bu kontrolü atla.
    # 'sys.frozen' özelliği sadece PyInstaller .exe'lerinde bulunur.
    if not getattr(sys, 'frozen', False):
        install_and_check()
    
    # Import deferred to avoid crashes if dependencies are missing during initial load
    from allone.app_ui import ToolApp
    
    # Uygulama arayüzünü başlat.
    app = ToolApp()
    app.mainloop()

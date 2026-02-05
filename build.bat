@echo off
echo ===========================================
echo       AllOne Tools - Auto Build Script
echo ===========================================
echo.
echo Installing requirements...
python -m pip install -r allone/requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install requirements.
    pause
    exit /b %errorlevel%
)

echo.
echo Building executable with PyInstaller...
python -m PyInstaller --noconfirm allone_onedir.spec
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed! Check the error log above.
    pause
    exit /b %errorlevel%
)

echo.
echo ===========================================
echo       Build Complete Successfully!
echo ===========================================
echo Output: dist\AllOne Tools\AllOne Tools.exe
echo.
pause

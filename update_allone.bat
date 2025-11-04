@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PYTHON_CMD="

for %%P in (python.exe py.exe) do (
    where %%P >NUL 2>&1
    if not errorlevel 1 (
        if /I "%%~nP"=="py" (
            set "PYTHON_CMD=py -3"
        ) else (
            set "PYTHON_CMD=python"
        )
        goto :found_python
    )
)

echo Python interpreter could not be found. Please install Python 3 and try again.
goto :eof

:found_python
pushd "%SCRIPT_DIR%"
%PYTHON_CMD% -m allone.update_cli --install-root "%SCRIPT_DIR%" --wait %*
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
    echo Update failed with exit code %EXIT_CODE%.
    pause
    exit /B %EXIT_CODE%
)

echo Update completed successfully.
pause
exit /B 0

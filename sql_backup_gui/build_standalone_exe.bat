@echo off
setlocal
cd /d "%~dp0"

echo ============================================================================
echo Building Standalone Windows Executable (dist\sql_backup_gui.exe)
echo ============================================================================

python -m PyInstaller --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install PyInstaller. Ensure pip and internet access are available.
        exit /b 1
    )
)

python -m PyInstaller --noconfirm --onefile --windowed --name "sql_backup_gui" "%~dp0sql_backup_gui.py"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build complete: "%~dp0dist\sql_backup_gui.exe"
) else (
    echo.
    echo Build failed.
    exit /b 1
)

@echo off
setlocal
cd /d "%~dp0"

echo ============================================================================
echo Building Standalone Windows Executable (dist\aws_cp_gui.exe)
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

python -m PyInstaller --noconfirm --onefile --windowed --name "aws_cp_gui" "%~dp0aws_cp_gui.py"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build complete: "%~dp0dist\aws_cp_gui.exe"
) else (
    echo.
    echo Build failed.
    exit /b 1
)

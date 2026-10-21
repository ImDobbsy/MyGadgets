@echo off
setlocal
cd /d "%~dp0"

:: Allow forcing a specific runtime via `--python` or `--powershell`
if /I "%~1"=="--powershell" goto :run_ps
if /I "%~1"=="--python" goto :run_py

:: Check if a compiled standalone executable exists in dist\ or current folder
if exist "%~dp0sql_restore_gui.exe" (
    start "" "%~dp0sql_restore_gui.exe"
    exit /b 0
)
if exist "%~dp0dist\sql_restore_gui.exe" (
    start "" "%~dp0dist\sql_restore_gui.exe"
    exit /b 0
)

:: Check if Python with Tkinter is available (User Desktops)
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    python -c "import tkinter" >nul 2>&1
    if %ERRORLEVEL% EQU 0 goto :run_py
)

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    py -3 -c "import tkinter" >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        start "" pyw -3 "%~dp0sql_restore_gui.py"
        exit /b 0
    )
)

:: Fallback for Windows Servers without Python installed -> Native PowerShell/.NET WinForms GUI
:run_ps
echo Launching Native Windows Server Edition (sql_restore_gui.ps1)...
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0sql_restore_gui.ps1"
exit /b %ERRORLEVEL%

:run_py
where pythonw >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    start "" pythonw "%~dp0sql_restore_gui.py"
) else (
    python "%~dp0sql_restore_gui.py"
)
exit /b %ERRORLEVEL%

@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"

set "APP_FILE=invoice_manager_gui.py"
set "APP_HOST=127.0.0.1"
set "APP_PORT=8000"
set "APP_URL=http://%APP_HOST%:%APP_PORT%"
set "PYTHON_CMD="
set "PYTHON_ARGS="
set "PYTHON_LABEL="

set "INSTALL_PACKAGES=flask openpyxl qrcode[pil] reportlab pypdf pypdfium2 rapidocr-onnxruntime"

if exist ".venv\Scripts\python.exe" (
    set "PYTHON_CMD=%CD%\.venv\Scripts\python.exe"
    set "PYTHON_LABEL=lokalni .venv"
)

if not defined PYTHON_CMD if exist "venv\Scripts\python.exe" (
    set "PYTHON_CMD=%CD%\venv\Scripts\python.exe"
    set "PYTHON_LABEL=lokalni venv"
)

if not defined PYTHON_CMD (
    where py >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py"
        set "PYTHON_ARGS=-3"
        set "PYTHON_LABEL=py -3"
    )
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=python"
        set "PYTHON_LABEL=python"
    )
)

if not defined PYTHON_CMD (
    echo Python nebyl nalezen.
    echo.
    echo Na tomto pocitaci chybi Python 3.
    echo Pro spusteni pres BAT je potreba nainstalovat Python a povolit prikaz python nebo py -3.
    echo.
    echo Doporuceni:
    echo 1. Nainstaluj Python 3 z https://www.python.org/downloads/
    echo 2. Pri instalaci zaskrtni Add Python to PATH
    echo 3. Spust znovu tento soubor
    echo.
    echo Alternativa: pouzij prenositelny EXE build aplikace.
    pause
    exit /b 1
)

echo Nalezen Python pres: %PYTHON_LABEL%
echo Kontroluji potrebne knihovny...

"%PYTHON_CMD%" %PYTHON_ARGS% -c "import importlib.util,sys;mods='flask openpyxl qrcode reportlab pypdf pypdfium2 rapidocr_onnxruntime'.split();missing=[m for m in mods if importlib.util.find_spec(m) is None];print(','.join(missing));sys.exit(1 if missing else 0)" > "%TEMP%\fakturace_missing_modules.txt"
set "MISSING_EXIT=%ERRORLEVEL%"
set "MISSING_MODULES="
if exist "%TEMP%\fakturace_missing_modules.txt" set /p MISSING_MODULES=<"%TEMP%\fakturace_missing_modules.txt"
del "%TEMP%\fakturace_missing_modules.txt" >nul 2>nul

if not "%MISSING_EXIT%"=="0" (
    echo.
    echo Chybi tyto knihovny:
    echo %MISSING_MODULES%
    echo.
    choice /M "Chces je ted doinstalovat automaticky"
    if errorlevel 2 (
        echo Instalace byla zrusena. Aplikace se nespusti.
        pause
        exit /b 1
    )

    echo.
    echo Pripravuji pip...
    "%PYTHON_CMD%" %PYTHON_ARGS% -m ensurepip --upgrade >nul 2>nul
    echo Instaluji: %INSTALL_PACKAGES%
    "%PYTHON_CMD%" %PYTHON_ARGS% -m pip install --upgrade pip
    if errorlevel 1 (
        echo Nepodarilo se aktualizovat pip.
        pause
        exit /b 1
    )
    "%PYTHON_CMD%" %PYTHON_ARGS% -m pip install %INSTALL_PACKAGES%
    if errorlevel 1 (
        echo Instalace knihoven selhala. Zkontroluj internet nebo prava k instalaci.
        pause
        exit /b 1
    )
    echo Knihovny byly uspesne doinstalovany.
)

echo.
echo Spoustim aplikaci Fakturace Studio...
echo Databaze SQLite se nacte automaticky pri startu aplikace.
echo Web se otevre za 2 sekundy na adrese %APP_URL%
echo.

echo Ukoncuji predchozi instance aplikace...
set "KILL_SCRIPT=%TEMP%\fakturace_kill_old.ps1"
(
  echo $ErrorActionPreference = 'SilentlyContinue'
  echo $procs = Get-CimInstance Win32_Process ^| Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -like '*invoice_manager_gui.py*' }
  echo foreach ^($p in $procs^) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
) > "%KILL_SCRIPT%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%KILL_SCRIPT%" >nul 2>nul
del "%KILL_SCRIPT%" >nul 2>nul

start "" /min powershell -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process '%APP_URL%'"

set "APP_HOST=%APP_HOST%"
set "APP_PORT=%APP_PORT%"
set "APP_PUBLIC_URL=%APP_URL%"

"%PYTHON_CMD%" %PYTHON_ARGS% "%APP_FILE%"

echo.
echo Aplikace byla ukoncena.
pause

@echo off
chcp 65001 >nul
title THNDR Portfolio Analyzer

echo.
echo ======================================================
echo          THNDR PORTFOLIO ANALYZER
echo          First run installs everything
echo ======================================================
echo.

:: ── PATHS ─────────────────────────────────────────────────
set BAT_DIR=%~dp0
set BAT_DIR=%BAT_DIR:~0,-1%
set SCRAPER=%BAT_DIR%\Scraper_app.py
set ANALYZER=%BAT_DIR%\full_app.py
set DESKTOP=%USERPROFILE%\Desktop

:: ── Build dated output folder on Desktop and use it as working dir ──
:: Use PowerShell for the date — %DATE% format varies by locale (some PCs
:: prefix a day name like "Mon 11/05/2026" which breaks the FOR /F parser).
:: PowerShell gives us a guaranteed YYYY-MM-DD regardless of regional settings.
set TODAY=
for /f "delims=" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set TODAY=%%d
if "%TODAY%"=="" (
    echo WARNING: PowerShell date failed, falling back to undated folder.
    set TODAY=undated
)
set OUTPUT=%DESKTOP%\Thndr_%TODAY%
if not exist "%OUTPUT%" mkdir "%OUTPUT%"

:: ==========================================================
:: STEP 1: Find Python — SKIP Windows Store alias
:: ==========================================================
echo [1/6] Looking for Python...
set PYTHON=

for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python38\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
    "C:\Python39\python.exe"
    "C:\Program Files\Python312\python.exe"
    "C:\Program Files\Python311\python.exe"
    "C:\Program Files\Python310\python.exe"
    "C:\Program Files (x86)\Python310\python.exe"
) do (
    if exist %%P (
        set PYTHON=%%P
        goto :found_python
    )
)

echo Python not found. Installing Python 3.10...
echo.

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\App Paths\python.exe" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" /v "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" /f >nul 2>&1

echo Downloading Python 3.10 installer...
curl -# -L -o "%TEMP%\python_installer.exe" "https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Download failed. Please check your internet and try again.
    pause ^& exit /b 1
)

echo Installing Python (2-3 minutes, a window may flash briefly)...
"%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=0
echo Waiting for install to complete...
timeout /t 20 >nul

set PYTHON=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
if not exist "%PYTHON%" (
    echo.
    echo ====================================================
    echo  Python installed but needs a fresh window.
    echo  Please CLOSE this window and run the bat again.
    echo ====================================================
    pause ^& exit /b 0
)

:found_python
echo Python: %PYTHON%
echo.

:: ==========================================================
:: STEP 2: Upgrade pip
:: ==========================================================
echo [2/6] Updating pip...
"%PYTHON%" -m pip install --upgrade pip --quiet
echo Done.
echo.

:: ==========================================================
:: STEP 3: Install all packages
:: ==========================================================
echo [3/6] Installing packages (3-5 min first time, instant after)...
"%PYTHON%" -m pip install --quiet pandas requests openai beautifulsoup4 openpyxl crawl4ai arabic-reshaper python-bidi fpdf2 python-dotenv playwright
if %ERRORLEVEL% neq 0 (
    echo ERROR: Package install failed. Check internet connection.
    pause ^& exit /b 1
)
echo Done.
echo.

:: ==========================================================
:: STEP 4: Install Chromium browser
:: ==========================================================
echo [4/6] Installing Chromium browser...
"%PYTHON%" -m playwright install chromium
if %ERRORLEVEL% neq 0 (
    echo ERROR: Chromium install failed.
    pause ^& exit /b 1
)
echo Done.
echo.

:: ==========================================================
:: STEP 5: Run scraper (working dir = Desktop output folder,
:: so the scraper writes its xlsx there, NOT next to the .bat)
:: ==========================================================
echo [5/6] Opening Thndr scraper...
echo.
echo ========================================================
echo  A browser will open. Log in to your Thndr account.
echo  Once you can SEE your investments on screen,
echo  come back here and press ENTER to continue.
echo.
echo  Output folder: %OUTPUT%
echo ========================================================
echo.
cd /d "%OUTPUT%"
"%PYTHON%" "%SCRAPER%"
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Scraper failed. Make sure you logged in correctly.
    pause ^& exit /b 1
)
echo.
echo Excel saved to: %OUTPUT%
echo.

:: ==========================================================
:: Verify the scraper actually produced an Excel file there
:: ==========================================================
dir /b "%OUTPUT%\investments_*.xlsx" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: No investments_*.xlsx file was found in:
    echo   %OUTPUT%
    echo.
    echo Make sure Scraper_app.py uses os.getcwd^(^) for SAVE_FOLDER.
    pause ^& exit /b 1
)

:: ==========================================================
:: STEP 6: Run AI analysis (still in Desktop output folder,
:: so it reads the xlsx from there and writes the report there)
:: ==========================================================
echo [6/6] Running AI analysis (1-3 minutes)...
echo.
cd /d "%OUTPUT%"
"%PYTHON%" "%ANALYZER%"
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Analysis failed. Check internet connection.
    pause ^& exit /b 1
)
echo.

:: ==========================================================
:: Done — files are already on the Desktop, no copying needed
:: ==========================================================
echo.
echo ======================================================
echo                   ALL DONE!
echo.
echo   Your files are on the Desktop in:
echo   Thndr_%TODAY%
echo.
echo   - investments_*.xlsx      (portfolio data)
echo   - portfolio_report_*.txt  (AI analysis)
echo ======================================================
echo.
explorer "%OUTPUT%"
pause
@echo off
REM ============================================================
REM ExcelRouter - one-click packaging script (Windows)
REM Copyright (c) 2026 AbeLin
REM ------------------------------------------------------------
REM NOTE: this file is intentionally ASCII-only. cmd.exe parses a
REM .bat using the console's active code page (936/GBK on Chinese
REM Windows). A UTF-8 (or UTF-8+BOM) file with CJK text gets mis-
REM parsed there and breaks the script. ASCII is byte-identical on
REM every code page, so it always parses. Keep new lines ASCII-only.
REM ------------------------------------------------------------
REM KEY: "--collect-all customtkinter" is mandatory. Without it the
REM packaged exe crashes on launch because it cannot find
REM customtkinter's theme/font assets (e.g. blue.json).
REM KEY: "--add-data config;config" is mandatory, otherwise the
REM default config is missing at runtime (Windows uses ';').
REM KEY: "--add-data app.ico;." is mandatory too. "--icon app.ico"
REM only stamps the exe FILE icon; the running window's title-bar/
REM taskbar icon is set at runtime via Tk iconbitmap(), which needs
REM app.ico bundled as data so gui/app.py can find it after packing.
REM KEY: "--collect-all tkinterdnd2" is mandatory since v2.7 (drag & drop).
REM Without it the tkdnd2.x.x binary is missing and drag-in falls back
REM to the silent no-dnd path.
REM
REM This script produces BOTH forms; prefer distributing onedir:
REM   - onedir : no self-extract at runtime, faster start, far less
REM             likely to trip antivirus false positives
REM   - onefile: single file, download-and-run, but some antivirus
REM             engines may false-positive (see docs/FAQ.md)
REM ============================================================

echo.
echo [1/4] Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/4] Cleaning previous build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "ExcelRouter.spec" del "ExcelRouter.spec"
if exist "ExcelRouter-onefile.spec" del "ExcelRouter-onefile.spec"

echo.
echo [3/4] Building onedir (recommended for distribution)...
pyinstaller --onedir --noconsole ^
  --name "ExcelRouter" ^
  --icon app.ico ^
  --version-file version.txt ^
  --collect-all customtkinter ^
  --collect-all tkinterdnd2 ^
  --add-data "config;config" ^
  --add-data "app.ico;." ^
  main.py

echo.
echo [4/4] Building onefile (alternative, single file)...
pyinstaller --onefile --noconsole ^
  --name "ExcelRouter-onefile" ^
  --icon app.ico ^
  --version-file version.txt ^
  --collect-all customtkinter ^
  --collect-all tkinterdnd2 ^
  --add-data "config;config" ^
  --add-data "app.ico;." ^
  main.py

echo.
if exist "dist\ExcelRouter\ExcelRouter.exe" if exist "dist\ExcelRouter-onefile.exe" (
    echo ============================================================
    echo  Build succeeded!
    echo   onedir  (recommended): dist\ExcelRouter\ExcelRouter.exe
    echo   onefile (alternative): dist\ExcelRouter-onefile.exe
    echo ============================================================
) else (
    echo  Build FAILED. Check the error messages above.
)
echo.
pause

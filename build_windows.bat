@echo off
REM ============================================================
REM  Portfolio Tracker v2 — Windows build script
REM  Produces:
REM    dist\PortfolioTracker\PortfolioTracker.exe  (portable bundle)
REM    installer\Output\PortfolioTracker-v2.0.0-Setup.exe  (installer)
REM
REM  Prerequisites:
REM    pip install pyinstaller pyinstaller-hooks-contrib
REM    Inno Setup 6.x  https://jrsoftware.org/isinfo.php
REM ============================================================

setlocal enabledelayedexpansion

echo ==========================================================
echo   Portfolio Tracker v2 — Windows Build
echo ==========================================================
echo.

REM ── Step 1: PyInstaller bundle ────────────────────────────
echo [1/3] Building PyInstaller bundle...
pyinstaller portfolio_tracker.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed. Check the output above.
    pause
    exit /b 1
)
echo       Done: dist\PortfolioTracker\PortfolioTracker.exe
echo.

REM ── Step 2: Optional code signing ─────────────────────────
REM Uncomment and set CERT_THUMBPRINT to sign the exe before packaging.
REM Requires Windows SDK signtool and a valid code-signing certificate.
REM
REM set CERT_THUMBPRINT=YOUR_CERT_THUMBPRINT_HERE
REM echo [2/3] Signing executable...
REM signtool sign /fd sha256 /sha1 %CERT_THUMBPRINT% /tr http://timestamp.sectigo.com /td sha256 dist\PortfolioTracker\PortfolioTracker.exe
REM if errorlevel 1 ( echo ERROR: Code signing failed. & pause & exit /b 1 )
REM echo       Done: executable signed.
REM echo.

echo [2/3] Code signing skipped (set CERT_THUMBPRINT to enable).
echo.

REM ── Step 3: Inno Setup installer ──────────────────────────
echo [3/3] Building installer with Inno Setup...

REM Locate ISCC.exe in common install paths
set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC=C:\Program Files\Inno Setup 6\ISCC.exe

if "!ISCC!"=="" (
    echo       WARNING: Inno Setup not found. Skipping installer creation.
    echo       Install from: https://jrsoftware.org/isinfo.php
    echo       Then re-run this script, or run manually:
    echo         "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
    echo.
) else (
    mkdir installer\Output 2>nul
    "!ISCC!" installer\setup.iss
    if errorlevel 1 (
        echo.
        echo ERROR: Inno Setup compilation failed. Check the output above.
        pause
        exit /b 1
    )
    echo       Done: installer\Output\PortfolioTracker-v2.0.0-Setup.exe
    echo.
)

REM ── Summary ───────────────────────────────────────────────
echo ==========================================================
echo   Build complete!
echo.
echo   Portable bundle : dist\PortfolioTracker\PortfolioTracker.exe
if not "!ISCC!"=="" (
echo   Installer       : installer\Output\PortfolioTracker-v2.0.0-Setup.exe
)
echo.
echo   Ship either file to end users.
echo   See INSTALL.md for distribution notes.
echo ==========================================================
pause

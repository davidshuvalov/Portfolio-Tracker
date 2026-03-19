@echo off
REM Portfolio Tracker v2 — Windows build script
REM Prerequisites: pip install pyinstaller pyinstaller-hooks-contrib

echo Building Portfolio Tracker v2...
pyinstaller portfolio_tracker.spec --clean --noconfirm

REM Code signing (requires Windows SDK signtool and a code-signing certificate)
REM Uncomment and set CERT_THUMBPRINT to enable:
REM set CERT_THUMBPRINT=YOUR_CERT_THUMBPRINT_HERE
REM signtool sign /fd sha256 /sha1 %CERT_THUMBPRINT% /tr http://timestamp.sectigo.com /td sha256 dist\PortfolioTracker\PortfolioTracker.exe

echo.
echo Build complete: dist\PortfolioTracker\PortfolioTracker.exe
pause

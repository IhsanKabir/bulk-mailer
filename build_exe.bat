@echo off
REM ================================================================
REM Build BulkMailer.exe with PyInstaller
REM ================================================================
REM Output: dist\BulkMailer.exe (~30-40 MB)
REM
REM End-user machines need:
REM   - NO Python install
REM   - NO admin access
REM   - Outlook desktop ONLY if using the Outlook transport
REM     (Microsoft 365 / SMTP transports work without it)
REM ================================================================

setlocal

echo [1/3] Installing build deps...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 goto :err

echo [2/3] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] Building one-file exe...
python -m PyInstaller --noconfirm BulkMailer.spec
if errorlevel 1 goto :err

echo.
echo ===============================================
echo  Build done.  dist\BulkMailer.exe
echo ===============================================
echo.
goto :eof

:err
echo.
echo *** Build failed. See errors above.
exit /b 1

@echo off
rem Bulk Mail Sender launcher (Windows)
set "PYTHONPATH=%~dp0libs"

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%~dp0mail_sender.py"
    exit /b
)

if exist "D:\python\pythonw.exe" (
    start "" "D:\python\pythonw.exe" "%~dp0mail_sender.py"
    exit /b
)

echo [Bulk Mail Sender] Python not found.
echo Please install Python 3.8+ from https://www.python.org/downloads/
echo and check "Add python.exe to PATH" during installation.
pause

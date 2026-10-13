@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Building Email Finder EXE for Windows
echo ==========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON=py"
) else (
    set "PYTHON=python"
)

%PYTHON% --version
if errorlevel 1 (
    echo ERROR: Python was not found. Install Python from https://www.python.org/downloads/windows/
    echo During installation, enable "Add Python to PATH".
    pause
    exit /b 1
)

echo.
echo Installing required packages...
%PYTHON% -m pip install --upgrade pip
if errorlevel 1 goto :failed

%PYTHON% -m pip install openpyxl pyinstaller
if errorlevel 1 goto :failed

echo.
echo Building the EXE...
%PYTHON% -m PyInstaller --clean --onefile --name email_finder email_finder.py
if errorlevel 1 goto :failed

echo.
echo SUCCESS! Your EXE is here:
echo %~dp0dist\email_finder.exe
echo.
pause
exit /b 0

:failed
echo.
echo BUILD FAILED. Review the error messages above.
pause
exit /b 1

@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ================================================
echo   BAP Donor Traceability - Windows EXE Builder
echo ================================================
echo.

where py >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python launcher ^(py^) was not found.
    echo Install Python 3.13 for Windows first.
    pause
    exit /b 1
)

if not exist "C:\msys64\ucrt64\bin" (
    echo ERROR: MSYS2 UCRT64 was not found.
    echo Expected folder: C:\msys64\ucrt64\bin
    echo.
    echo If that folder exists, verify that this exact path is present.
    pause
    exit /b 1
)

set "VENV=.build-venv"
if not exist "%VENV%\Scripts\python.exe" (
    echo Creating build environment...
    py -3.13 -m venv "%VENV%"
    if errorlevel 1 goto :fail
)

call "%VENV%\Scripts\activate.bat"
if errorlevel 1 goto :fail

python -m pip install --upgrade pip
if errorlevel 1 goto :fail

python -m pip install -r requirements-weasy.txt pyinstaller pefile
if errorlevel 1 goto :fail

set "WEASYPRINT_DLL_DIRECTORIES=C:\msys64\ucrt64\bin"

if exist msys_dlls.txt del /q msys_dlls.txt
python collect_msys_dlls.py
if errorlevel 1 goto :fail

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

python -m PyInstaller --clean --noconfirm BAP_Donor_Report.spec
if errorlevel 1 goto :fail

if not exist "dist\BAP_Donor_Traceability.exe" goto :fail

echo.
echo ================================================
echo   BUILD COMPLETE
echo ================================================
echo.
echo Final single-file application:
echo   %CD%\dist\BAP_Donor_Traceability.exe
echo.
echo IMPORTANT: Test this EXE before sending it to the recipient.
echo.
pause
exit /b 0

:fail
echo.
echo BUILD FAILED.
echo Review the error above.
pause
exit /b 1

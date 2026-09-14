@echo off
setlocal

python -m pip install -U pyinstaller pandas openpyxl
if errorlevel 1 (
    echo Failed to install dependencies.
    pause
    exit /b 1
)

pyinstaller --noconfirm --clean --onefile --windowed --name DonorTraceabilityGUI report_gui.py

if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete: dist\DonorTraceabilityGUI.exe
pause

@echo off
cd /d "%~dp0"
if not exist .build-venv\Scripts\python.exe py -3.13 -m venv .build-venv
call .build-venv\Scripts\activate.bat
python -m pip install -r requirements.txt
set WEASYPRINT_DLL_DIRECTORIES=C:\msys64\ucrt64\bin
python app.py
pause

@echo off
setlocal EnableExtensions
cd /d "%~dp0"
echo =======================================================
echo   BAP Report Generator V3.0 - Exact Money Template Build
echo =======================================================
echo.
where py >nul 2>nul
if errorlevel 1 (
 echo ERROR: No se encontro el comando py.
 echo Instale Python 3.13 para Windows y vuelva a intentar.
 goto :fail
)
py -3.13 --version >nul 2>nul
if errorlevel 1 (
 echo ERROR: Este build usa Python 3.13 para mantener compatibilidad con WeasyPrint.
 echo Instale Python 3.13 x64 y vuelva a ejecutar este archivo.
 goto :fail
)
if not exist "C:\msys64\ucrt64\bin" (
 echo ERROR: No se encontro MSYS2 UCRT64 en C:\msys64\ucrt64\bin
 echo Cam utiliza estas librerias para empaquetar WeasyPrint en Windows.
 goto :fail
)
set "VENV=.build-venv"
if not exist "%VENV%\Scripts\python.exe" py -3.13 -m venv "%VENV%"
if errorlevel 1 goto :fail
call "%VENV%\Scripts\activate.bat"
python -m pip install --upgrade pip
if errorlevel 1 goto :fail
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail
set "WEASYPRINT_DLL_DIRECTORIES=C:\msys64\ucrt64\bin"
if exist msys_dlls.txt del /q msys_dlls.txt
python collect_msys_dlls.py
if errorlevel 1 goto :fail
set "PLAYWRIGHT_BROWSERS_PATH=%CD%\ms-playwright"
python -m playwright install chromium
if errorlevel 1 goto :fail
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
python -m PyInstaller --clean --noconfirm BAP_Integration.spec
if errorlevel 1 goto :fail
if not exist "dist\Generador de Reportes BAP\Generador de Reportes BAP.exe" goto :fail
echo Copiando plantillas editables y manual...
if exist "dist\Generador de Reportes BAP\templates\money" rmdir /s /q "dist\Generador de Reportes BAP\templates\money"
xcopy /E /I /Y "templates\money" "dist\Generador de Reportes BAP\templates\money" >nul
copy /Y "Guia_Generador_Informes_Trazabilidad_BAP_Espanol.pdf" "dist\Generador de Reportes BAP\" >nul
echo.
echo =======================================================
echo BUILD COMPLETO
echo =======================================================
echo La aplicacion esta en:
echo dist\Generador de Reportes BAP\Generador de Reportes BAP.exe
echo.
echo IMPORTANTE: entregue/pruebe la carpeta completa, no solo el EXE.
pause
exit /b 0
:fail
echo.
echo BUILD FALLIDO. Revise el primer mensaje de ERROR arriba.
echo Si necesita ayuda, envie una captura de las ultimas 30 lineas.
pause
exit /b 1

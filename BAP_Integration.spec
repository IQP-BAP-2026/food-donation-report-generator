from pathlib import Path
from PyInstaller.utils.hooks import collect_all
ROOT=Path(SPECPATH)
weasy_datas,weasy_binaries,weasy_hidden=collect_all('weasyprint')
dll_list=ROOT/'msys_dlls.txt'
if not dll_list.exists(): raise SystemExit('msys_dlls.txt no existe. Ejecute BUILD_EXE.bat.')
for line in dll_list.read_text(encoding='utf-8').splitlines():
    line=line.strip()
    if line: weasy_binaries.append((line,'weasy_dlls'))
datas=[
 (str(ROOT/'templates'),'templates'),(str(ROOT/'assets'),'assets'),
 (str(ROOT/'food'),'food'),
(str(ROOT/'Guia_Generador_Informes_Trazabilidad_BAP_Espanol.pdf'),'.'),(str(ROOT/'ms-playwright'),'ms-playwright'),*weasy_datas]
a=Analysis([str(ROOT/'app.py')],pathex=[str(ROOT)],binaries=weasy_binaries,datas=datas,
 hiddenimports=['food_generator','engine.food_adapter','openpyxl','pandas','jinja2','weasyprint','playwright','PIL','PIL.Image','PIL.ImageOps',*weasy_hidden],
 excludes=['matplotlib','scipy','pytest'],noarchive=False)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,[],exclude_binaries=True,name='Generador de Reportes BAP',debug=False,bootloader_ignore_signals=False,strip=False,upx=False,console=False,icon=str(ROOT/'bap_icon.ico'))
coll=COLLECT(exe,a.binaries,a.datas,strip=False,upx=False,name='Generador de Reportes BAP')

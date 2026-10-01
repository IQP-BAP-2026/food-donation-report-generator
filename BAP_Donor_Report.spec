from pathlib import Path
from PyInstaller.utils.hooks import collect_all

ROOT = Path(SPECPATH)

weasy_datas, weasy_binaries, weasy_hidden = collect_all("weasyprint")

# The generator loads the Windows GTK/Pango stack through WeasyPrint.
# build_windows.bat creates this list by following the actual PE imports,
# instead of blindly bundling every MSYS2 DLL.
dll_list = ROOT / "msys_dlls.txt"
if not dll_list.exists():
    raise SystemExit("msys_dlls.txt was not generated. Run build_windows.bat again.")

for line in dll_list.read_text(encoding="utf-8").splitlines():
    line = line.strip()
    if line:
        weasy_binaries.append((line, "weasy_dlls"))

analysis = Analysis(
    [str(ROOT / "app.py")],
    pathex=[str(ROOT)],
    binaries=weasy_binaries,
    datas=[
        (str(ROOT / "food-traceability-template.html"), "."),
        (str(ROOT / "food-traceability.css"), "."),
        (str(ROOT / "MASTER-SHEET.xlsx"), "."),
        (str(ROOT / "assets"), "assets"),
        *weasy_datas,
    ],
    hiddenimports=[
        "food_generator",
        *weasy_hidden,
        "openpyxl",
        "pandas",
        "jinja2",
        "weasyprint",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "pytest"],
    noarchive=False,
)

pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="BAP_Donor_Traceability",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon='bap_icon.ico'
)

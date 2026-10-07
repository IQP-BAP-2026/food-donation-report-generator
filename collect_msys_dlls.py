from pathlib import Path
import re
import pefile

BIN = Path(r"C:\msys64\ucrt64\bin")
OUT = Path("msys_dlls.txt")

# Native libraries used by WeasyPrint/Pango/Cairo on Windows.
ROOTS = [
    "libgobject-2.0-0.dll",
    "libpango-1.0-0.dll",
    "libpangocairo-1.0-0.dll",
    "libcairo-2.dll",
]

seen = set()
queue = list(ROOTS)
found = []

# DLLs supplied by Windows should not be bundled.
SYSTEM_DLLS = {
    "kernel32.dll", "kernelbase.dll", "user32.dll", "gdi32.dll", "gdi32full.dll",
    "advapi32.dll", "ole32.dll", "oleaut32.dll", "shell32.dll", "shlwapi.dll",
    "ws2_32.dll", "secur32.dll", "bcrypt.dll", "crypt32.dll", "combase.dll",
    "ntdll.dll", "rpcrt4.dll", "msvcrt.dll", "ucrtbase.dll", "version.dll",
    "imm32.dll", "winmm.dll", "winspool.drv", "comdlg32.dll", "dwmapi.dll",
    "userenv.dll", "psapi.dll", "dbghelp.dll", "setupapi.dll", "cfgmgr32.dll",
    "opengl32.dll", "glu32.dll", "avrt.dll", "dnsapi.dll", "iphlpapi.dll",
    "normaliz.dll", "ncrypt.dll", "bcryptprimitives.dll", "mswsock.dll",
    "winhttp.dll", "wtsapi32.dll", "netapi32.dll", "version.dll",
}

while queue:
    name = queue.pop(0).lower()
    if name in seen or name in SYSTEM_DLLS:
        continue
    seen.add(name)
    path = BIN / name
    if not path.exists():
        # Case-insensitive fallback for MSYS2 filenames.
        matches = list(BIN.glob(name))
        if not matches:
            matches = [p for p in BIN.iterdir() if p.is_file() and p.name.lower() == name]
        if not matches:
            print(f"WARNING: missing dependency {name}")
            continue
        path = matches[0]
    found.append(path)
    try:
        pe = pefile.PE(str(path), fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_IMPORT']])
        for entry in getattr(pe, "DIRECTORY_ENTRY_IMPORT", []):
            dep = entry.dll.decode("ascii", errors="ignore").lower()
            if dep not in seen and dep not in SYSTEM_DLLS:
                queue.append(dep)
    except Exception as exc:
        print(f"WARNING: could not inspect {path.name}: {exc}")

# Also include common runtime DLLs when present; they are frequently loaded
# indirectly by the native stack and are safe to ship with the app.
for extra in ["libgcc_s_seh-1.dll", "libstdc++-6.dll", "libwinpthread-1.dll"]:
    p = BIN / extra
    if p.exists() and p not in found:
        found.append(p)

with OUT.open("w", encoding="utf-8") as f:
    for p in sorted(found, key=lambda x: x.name.lower()):
        f.write(str(p) + "\n")

print(f"Collected {len(found)} native DLLs")
for p in sorted(found, key=lambda x: x.name.lower()):
    print(f"  {p.name}")

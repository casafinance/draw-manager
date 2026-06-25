# -*- mode: python ; coding: utf-8 -*-
# Build with:  pyinstaller --clean "casa-updater.spec"
#
# Casa Balance Updater — second windowed app, launched as a card from
# Draw Manager. Bundles the HTML UI + all of Casa's logic deps (selenium for
# Chrome attach, PyMuPDF/pytesseract/Pillow for PDF OCR, Google libs for the
# Sheets API path). Entry is casa_main.py (pywebview window + CasaApi).
#
# NOTE: PDF OCR also needs the external Tesseract *engine* installed on the
# machine — that's a separate program, not bundleable here.

from PyInstaller.utils.hooks import collect_all

block_cipher = None

datas, binaries, hiddenimports = [], [], []


def _collect(mod):
    global datas, binaries, hiddenimports
    d, b, h = collect_all(mod)
    datas += d
    binaries += b
    hiddenimports += h


# pywebview (GUI shell)
_collect("webview")

# Google libs (Sheets API path)
for mod in ("googleapiclient", "google.auth", "google.oauth2",
            "google.auth.transport.requests", "google_auth_oauthlib",
            "google_auth_httplib2"):
    try:
        _collect(mod)
    except Exception:
        pass

# Excel
_collect("openpyxl")

# Selenium (Chrome remote-debug attach for FCI + view-only Sheets)
try:
    _collect("selenium")
except Exception:
    pass

# PDF + OCR stack
for mod in ("fitz", "pytesseract", "PIL"):
    try:
        _collect(mod)
    except Exception:
        pass

# Our own bundled assets
datas += [
    ("casa_updater.html", "."),
    ("draw_manager.ico",  "."),
]

# pythonnet bits pywebview needs explicit on Windows
hiddenimports += ["clr", "clr_loader", "clr_loader.netfx",
                  "bottle", "proxy_tools", "requests"]

a = Analysis(
    ["casa_main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["playwright"],   # Casa uses selenium, not playwright
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="casa-updater",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # windowed app, no console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="draw_manager.ico",
)

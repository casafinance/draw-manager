# -*- mode: python ; coding: utf-8 -*-
# Build with:  pyinstaller --clean "Draw Manager.spec"

from PyInstaller.utils.hooks import collect_all

block_cipher = None

datas, binaries, hiddenimports = [], [], []

# pywebview
d, b, h = collect_all("webview")
datas += d; binaries += b; hiddenimports += h

# Google libs
for mod in ("googleapiclient", "google.auth", "google.oauth2",
            "google.auth.transport.requests", "google_auth_httplib2"):
    d, b, h = collect_all(mod)
    datas += d; binaries += b; hiddenimports += h

# Local .xlsx reading
d, b, h = collect_all("openpyxl")
datas += d; binaries += b; hiddenimports += h

# Our own bundled assets
datas += [
    ("draw_manager.html", "."),
    ("draw_manager.ico",  "."),
]

# pythonnet needs clr_loader bits explicit
hiddenimports += ["clr", "clr_loader", "clr_loader.netfx",
                  "bottle", "proxy_tools"]

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["playwright"],   # the worker exe has Playwright; main doesn't need it
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
    name="Draw Manager",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="draw_manager.ico",
)

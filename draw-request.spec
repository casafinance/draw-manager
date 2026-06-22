# -*- mode: python ; coding: utf-8 -*-
# Build with:  pyinstaller --clean "draw-request.spec"

from PyInstaller.utils.hooks import collect_all

block_cipher = None

datas, binaries, hiddenimports = [], [], []

# Playwright + its driver / browser-locator files
d, b, h = collect_all("playwright")
datas += d; binaries += b; hiddenimports += h

# greenlet is a transitive dep playwright uses
hiddenimports += ["greenlet"]

a = Analysis(
    ["draw_request.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
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
    name="draw-request",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,                   # console so stdout streams back to parent
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

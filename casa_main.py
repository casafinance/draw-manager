"""
casa_main.py — entrypoint for the Casa Balance Updater window.

Creates a pywebview window loading casa_updater.html, backed by the CasaApi
bridge. This is the equivalent of Draw Manager's app.py main(): the GUI is
HTML, the logic is Python (casa_logic via casa_api).

Run standalone:  python casa_main.py
Bundled:         this is the entry script for casa-updater.exe (Stage 2).
"""

import os
import sys
from pathlib import Path

import webview

from casa_api import CasaApi


def _html_path():
    """Locate casa_updater.html next to this script, whether run from source
    or from a PyInstaller bundle (sys._MEIPASS)."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        p = Path(base) / "casa_updater.html"
        if p.exists():
            return str(p)
    here = Path(__file__).resolve().parent / "casa_updater.html"
    return str(here)


def main():
    api = CasaApi()
    window = webview.create_window(
        "Casa Balance Updater",
        url=_html_path(),
        js_api=api,
        width=1180,
        height=820,
        min_size=(960, 680),
    )
    api.set_window(window)
    webview.start()


if __name__ == "__main__":
    main()

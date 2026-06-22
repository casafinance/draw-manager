# Building Draw Manager.exe

## Prereqs

You should already have `.venv` set up from running `start.bat` once. If not, do that first.

## Build

Double-click `build.bat`. Takes ~2-3 minutes the first time.

Output ends up in `dist\`:
- `Draw Manager.exe` — the windowed app (no Python logo, no terminal, real icon)
- `draw-request.exe` — headless worker the app spawns to run the ProxyPics automation
- `draw_manager.ico` — copied alongside in case you want it for a desktop shortcut
- `_internal\` — PyInstaller's bundled-runtime folder (don't touch)

## Run

Just double-click `dist\Draw Manager.exe`.

Right-click it → Pin to taskbar / desktop to make it permanent.

On first launch the exe creates `settings.json`, `draws.db`, and `.pw-profile\` next to itself.

## When you change code

Re-run `build.bat`. It cleans the old `dist\` and rebuilds. Roughly 30-60 seconds for an incremental rebuild after the first cold build.

## Caveats

- **Playwright Chromium is NOT bundled** in the exe (would add ~300 MB). The worker uses your existing `%LOCALAPPDATA%\ms-playwright\` install. As long as `playwright install chromium` has been run on this machine, the exe will find it.
- **Antivirus** sometimes flags PyInstaller exes. If Defender quarantines `dist\draw-request.exe` on first launch, add the `dist\` folder as an exception, or sign the exe with a code-signing cert later if you want to distribute it.
- **First launch is slower** than subsequent ones — PyInstaller `--onefile` unpacks to a temp dir on first run.

## If something fails

The most common failure is missing Playwright internals. If `draw-request.exe` errors out with a `ModuleNotFoundError` or "driver not found", let me know and I'll add the missing pieces to `draw-request.spec`.

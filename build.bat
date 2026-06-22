@echo off
REM ---------------------------------------------------------------------------
REM Build Draw Manager into two .exe files:
REM   dist\Draw Manager.exe   — the windowed UI (with custom icon)
REM   dist\draw-request.exe   — headless automation worker
REM
REM Prereqs: run start.bat at least once so .venv exists with all deps.
REM ---------------------------------------------------------------------------

setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No .venv found. Run start.bat first.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

echo [build] Ensuring PyInstaller is installed...
.venv\Scripts\python.exe -c "import PyInstaller" 2>nul
if errorlevel 1 (
    python -m pip install pyinstaller || goto :error
)

echo [build] Cleaning old build artifacts...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo.
echo [build] (1/2) Building Draw Manager.exe (windowed UI)...
echo.
pyinstaller --clean --noconfirm "Draw Manager.spec" || goto :error

echo.
echo [build] (2/2) Building draw-request.exe (automation worker)...
echo.
pyinstaller --clean --noconfirm "draw-request.spec" || goto :error

echo.
echo [build] Consolidating outputs into dist\...
REM PyInstaller already drops both into dist\ — nothing to move.
REM We just need to copy supporting files for the user-facing folder.
copy /Y "draw_manager.ico" "dist\" >nul
if exist "SHEETS_SETUP.md" copy /Y "SHEETS_SETUP.md" "dist\" >nul

echo.
echo ============================================================
echo [build] DONE.
echo.
echo Open:  dist\Draw Manager.exe
echo.
echo First launch will create settings.json, draws.db, and
echo a .pw-profile folder right next to the exe.
echo ============================================================
echo.
pause
exit /b 0

:error
echo.
echo ============================================================
echo [BUILD FAILED] See messages above.
echo ============================================================
pause
exit /b 1

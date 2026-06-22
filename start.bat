@echo off
REM ---------------------------------------------------------------------------
REM Launch the Draw Manager GUI.
REM First run: creates .venv and installs everything (~1 minute).
REM ---------------------------------------------------------------------------

setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [setup] Creating virtual environment...
    python -m venv .venv || goto :error
    call .venv\Scripts\activate.bat
    python -m pip install --upgrade pip
    python -m pip install playwright pywebview google-api-python-client google-auth openpyxl || goto :error
    python -m playwright install chromium || goto :error
) else (
    call .venv\Scripts\activate.bat
    REM Top up any missing python deps from earlier installs.
    .venv\Scripts\python.exe -c "import webview, googleapiclient, google.oauth2, openpyxl, playwright" 2>nul
    if errorlevel 1 (
        echo [setup] Installing missing dependencies...
        python -m pip install pywebview google-api-python-client google-auth openpyxl playwright || goto :error
    )
    REM Make sure the chromium browser is present for draw_request.py.
    REM `playwright install chromium` is idempotent and only downloads if missing.
    .venv\Scripts\python.exe -m playwright install chromium 1>nul 2>nul
    if errorlevel 1 (
        echo [setup] Installing Playwright Chromium browser...
        python -m playwright install chromium || goto :error
    )
)

echo [run] Starting Draw Manager...
python app.py
goto :end

:error
echo.
echo [ERROR] Setup failed. See messages above.
pause
exit /b 1

:end
endlocal
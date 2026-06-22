@echo off
REM ===========================================================================
REM sync.bat - cut a new Draw Manager release.
REM
REM Reads the current VERSION from app.py, proposes the next patch version,
REM lets you accept it (Enter or y) or type a different one, validates that the
REM new version is strictly greater, then bumps app.py, commits, and pushes.
REM The push triggers the GitHub Actions build -> release pipeline.
REM ===========================================================================
setlocal enableextensions enabledelayedexpansion
cd /d "%~dp0"

REM --- make sure we're in a git repo with app.py present ---
if not exist "app.py" (
    echo [ERROR] app.py not found in this folder.
    goto :fail
)
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo [ERROR] This folder is not a git repository.
    goto :fail
)

REM --- read current VERSION from app.py. Inlining PowerShell inside for/f
REM     backticks collides with batch quote parsing, so instead PowerShell
REM     writes the version to a temp file and we read that plainly. ---
set "CURRENT="
set "_VTMP=%TEMP%\dm_ver_%RANDOM%.txt"
powershell -NoProfile -Command "$m = Select-String -Path 'app.py' -Pattern '^VERSION = \"([0-9]+\.[0-9]+\.[0-9]+)\"' | Select-Object -First 1; if ($m) { [System.IO.File]::WriteAllText('%_VTMP%', $m.Matches[0].Groups[1].Value) }"
if exist "%_VTMP%" (
    set /p CURRENT=<"%_VTMP%"
    del "%_VTMP%" >nul 2>&1
)

if "!CURRENT!"=="" (
    echo [ERROR] Could not read VERSION from app.py.
    goto :fail
)

REM --- split current into major.minor.patch ---
for /f "tokens=1-3 delims=." %%a in ("!CURRENT!") do (
    set "CMAJ=%%a" & set "CMIN=%%b" & set "CPAT=%%c"
)
REM --- propose next patch ---
set /a "NPAT=CPAT+1"
set "SUGGEST=!CMAJ!.!CMIN!.!NPAT!"

echo.
echo   Current version : !CURRENT!
echo   Suggested next  : !SUGGEST!
echo.
set "ANSWER="
set /p "ANSWER=Push new version  !CURRENT! -^> !SUGGEST!  ?  [Y/n, or type a version]: "

REM --- decide target version ---
if "!ANSWER!"=="" set "TARGET=!SUGGEST!"& goto :have_target
if /i "!ANSWER!"=="y" set "TARGET=!SUGGEST!"& goto :have_target
if /i "!ANSWER!"=="n" echo Cancelled.& goto :done
REM otherwise treat whatever they typed as a version string
set "TARGET=!ANSWER!"

:have_target
REM --- validate TARGET is x.y.z numeric ---
for /f "tokens=1-4 delims=." %%a in ("!TARGET!") do (
    set "TMAJ=%%a" & set "TMIN=%%b" & set "TPAT=%%c" & set "TEXTRA=%%d"
)
if not "!TEXTRA!"=="" (
    echo [ERROR] "!TARGET!" is not a valid x.y.z version.
    goto :fail
)
call :is_number "!TMAJ!" || (echo [ERROR] "!TARGET!" is not numeric x.y.z.& goto :fail)
call :is_number "!TMIN!" || (echo [ERROR] "!TARGET!" is not numeric x.y.z.& goto :fail)
call :is_number "!TPAT!" || (echo [ERROR] "!TARGET!" is not numeric x.y.z.& goto :fail)

REM --- ensure TARGET > CURRENT (numeric compare, major then minor then patch) ---
call :cmp_ver !TMAJ! !TMIN! !TPAT! !CMAJ! !CMIN! !CPAT!
if !CMPRESULT! LEQ 0 (
    echo [ERROR] New version !TARGET! must be GREATER than current !CURRENT!.
    goto :fail
)

echo.
echo   Bumping app.py  !CURRENT!  -^>  !TARGET!
echo.

REM --- rewrite the VERSION line in app.py via PowerShell (preserves the rest) ---
REM Pass TARGET through an env var so PowerShell reads it cleanly (no nested
REM batch !expansion! inside the quoted PS string, which breaks parsing).
set "NEWVER=!TARGET!"
powershell -NoProfile -Command "$p='app.py'; $v=$env:NEWVER; $c=Get-Content -Raw $p; $c=[regex]::Replace($c, '(?m)^VERSION = \"[0-9]+\.[0-9]+\.[0-9]+\"', ('VERSION = \"' + $v + '\"')); [System.IO.File]::WriteAllText((Resolve-Path $p), $c)"
if errorlevel 1 (
    echo [ERROR] Failed to update app.py.
    goto :fail
)

REM --- confirm it changed ---
findstr /r /c:"^VERSION = \"!TARGET!\"" app.py >nul
if errorlevel 1 (
    echo [ERROR] app.py was not updated to !TARGET! - aborting before commit.
    goto :fail
)

REM --- stage everything, commit, push ---
git add -A || goto :fail
git commit -m "Release v!TARGET!" || goto :fail
git push || goto :fail

echo.
echo ============================================================
echo   Pushed v!TARGET!.
echo   GitHub Actions is now building the release.
echo   Watch: https://github.com/casafinance/draw-manager/actions
echo   Release will appear at: .../releases/tag/v!TARGET!
echo ============================================================
echo.
goto :done

REM ---------------------------------------------------------------------------
:is_number
REM returns errorlevel 0 if %~1 is all digits, else 1
echo %~1| findstr /r "^[0-9][0-9]*$" >nul
exit /b %errorlevel%

:cmp_ver
REM args: aMaj aMin aPat bMaj bMin bPat ; sets CMPRESULT = 1 if A>B, 0 if equal, -1 if A<B
setlocal
set /a aMaj=%1, aMin=%2, aPat=%3, bMaj=%4, bMin=%5, bPat=%6
set "R=0"
if !aMaj! GTR !bMaj! set "R=1"& goto :cmp_done
if !aMaj! LSS !bMaj! set "R=-1"& goto :cmp_done
if !aMin! GTR !bMin! set "R=1"& goto :cmp_done
if !aMin! LSS !bMin! set "R=-1"& goto :cmp_done
if !aPat! GTR !bPat! set "R=1"& goto :cmp_done
if !aPat! LSS !bPat! set "R=-1"& goto :cmp_done
:cmp_done
endlocal & set "CMPRESULT=%R%"
exit /b 0

:fail
echo.
echo [ABORTED] No changes pushed.
echo.

:done
pause
endlocal
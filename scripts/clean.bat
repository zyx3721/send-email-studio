@echo off
rem ============================================================
rem  Launcher for scripts\clean.ps1
rem
rem  Double-click this file, or run from a terminal:
rem      clean.bat              dry run (lists what would be deleted)
rem      clean.bat execute      actually delete
rem      clean.bat purge        dry run, wipe the work folder too
rem      clean.bat execute purge   delete everything, including work
rem
rem  The work folder itself is kept by default: it may hold hand-written
rem  local verification scripts that git does not track. Only the leftover
rem  session/log files inside it are cleaned. -IncludeWork opts into the
rem  full wipe.
rem
rem  NOTE: this file is intentionally ASCII-only. cmd.exe mis-parses
rem  batch files that contain UTF-8 multi-byte characters, so all
rem  Chinese messages live in clean.ps1 instead.
rem
rem  NOTE: clean.ps1 must stay saved as "UTF-8 with BOM". Windows
rem  PowerShell 5.1 decodes a BOM-less .ps1 with the ANSI codepage, which
rem  turns the Chinese messages into mojibake and breaks parsing. The
rem  pre-flight check below turns that into one clear error message.
rem ============================================================

setlocal
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%clean.ps1"

if not exist "%PS1%" goto :missing

set "ARGS="
if /i "%~1"=="execute" set "ARGS=%ARGS% -Execute"
if /i "%~2"=="execute" set "ARGS=%ARGS% -Execute"
if /i "%~1"=="purge" set "ARGS=%ARGS% -IncludeWork"
if /i "%~2"=="purge" set "ARGS=%ARGS% -IncludeWork"

rem Pre-flight: refuse to run while clean.ps1 lacks its UTF-8 BOM.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$b=[IO.File]::ReadAllBytes('%PS1%');if($b.Length -lt 3 -or $b[0] -ne 0xEF -or $b[1] -ne 0xBB -or $b[2] -ne 0xBF){exit 9}"
if "%ERRORLEVEL%"=="9" goto :noBom

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %ARGS%
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" echo [done] Press any key to close.
if not "%EXITCODE%"=="0" echo [failed] See the messages above. Press any key to close.
pause >nul
endlocal
exit /b %EXITCODE%

:missing
echo [error] Not found: "%PS1%"
pause >nul
endlocal
exit /b 1

:noBom
echo.
echo [error] clean.ps1 is missing its UTF-8 byte order mark (BOM).
echo         Windows PowerShell 5.1 reads a BOM-less script with the ANSI
echo         codepage, so the Chinese messages turn into mojibake and the
echo         script fails to parse. Re-save the file as "UTF-8 with BOM",
echo         or run this once:
echo.
echo   powershell -NoProfile -Command "$p='%PS1%';[IO.File]::WriteAllText($p,(Get-Content -LiteralPath $p -Raw -Encoding UTF8),(New-Object Text.UTF8Encoding $true))"
echo.
pause >nul
endlocal
exit /b 1

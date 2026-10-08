@echo off
rem ============================================================
rem  Launcher for scripts\clean.ps1
rem
rem  Double-click this file, or run from a terminal:
rem      clean.bat              dry run (lists what would be deleted)
rem      clean.bat execute      actually delete
rem      clean.bat execute keep delete but keep the work folder
rem      clean.bat keep         dry run, keep the work folder
rem
rem  NOTE: this file is intentionally ASCII-only. cmd.exe mis-parses
rem  batch files that contain UTF-8 multi-byte characters, so all
rem  Chinese messages live in clean.ps1 instead.
rem ============================================================

setlocal
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%clean.ps1"

if not exist "%PS1%" goto :missing

set "ARGS="
if /i "%~1"=="execute" set "ARGS=-Execute"
if /i "%~1"=="keep" set "ARGS=-KeepWork"
if /i "%~2"=="keep" set "ARGS=%ARGS% -KeepWork"

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

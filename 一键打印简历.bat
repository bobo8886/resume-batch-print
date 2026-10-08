@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo [ERROR] The print tool exited with code %RC%.
  echo Press any key to close this window.
  pause >nul
)
endlocal

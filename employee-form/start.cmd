@echo off
REM ===================================================================
REM  Employee Form -> Excel   (double-click launcher for Windows)
REM  Prefers PowerShell 7 (pwsh), falls back to Windows PowerShell 5.1.
REM  server.ps1 installs its own dependencies and picks a free port.
REM ===================================================================
cd /d "%~dp0"

where pwsh >nul 2>&1
if %errorlevel%==0 (
    pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" %*
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" %*
)

pause

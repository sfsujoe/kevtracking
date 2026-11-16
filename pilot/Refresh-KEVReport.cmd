@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0scripts\Refresh-KEVReport.ps1" %*
set "refreshExitCode=%errorlevel%"
exit /b %refreshExitCode%

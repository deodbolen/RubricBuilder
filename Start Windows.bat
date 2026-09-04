@echo off
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js 20 or newer is required to run Rubric Builder.
  echo Install the LTS version from https://nodejs.org/, then run this starter again.
  pause
  exit /b 1
)
node scripts\start-local.mjs
echo.
echo Rubric Builder stopped. You can close this window.
pause

@echo off
title Excel Node App Server (port 3500)
cd /d "%~dp0"
set PORT=3500

where node >nul 2>nul
if errorlevel 1 (
  echo [X] Node.js not found in PATH. Please install Node.js 18+ from https://nodejs.org
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo [i] node_modules not found - running npm install first...
  call npm install
  if errorlevel 1 (
    echo [X] npm install failed.
    pause
    exit /b 1
  )
)

echo Starting server on http://localhost:3500 ...
start "" /min cmd /c "timeout /t 2 /nobreak >nul & start "" http://localhost:3500"
node server.js
echo.
echo Server stopped.
pause

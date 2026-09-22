@echo off
title Excel Node App Server (port 3000)
cd /d "%~dp0"
start "" /min cmd /c "timeout /t 2 /nobreak >nul & start "" http://localhost:3000"
node server.js
echo.
echo Server stopped.
pause

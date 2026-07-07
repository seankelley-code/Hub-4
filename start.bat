@echo off
cd /d "%~dp0"
echo.
echo  Starting Audit AI...
echo  Open http://localhost:3003 in your browser
echo  Press Ctrl+C to stop
echo.
node server.js
pause

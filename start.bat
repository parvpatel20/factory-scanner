@echo off
echo ==========================================
echo   Factory Scanner
echo ==========================================
echo.
echo Starting server... Open http://localhost:5050 in your browser
echo Press Ctrl+C to stop.
echo.
set PORT=5050
python server.py
pause

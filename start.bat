@echo off
chcp 65001 >nul
title Generator BR - Borderou de Reconsemnare
echo.
echo ============================================
echo    Generator BR - Borderou de Reconsemnare
echo ============================================
echo.

cd /d "%~dp0"

pip install -r requirements.txt --quiet 2>nul

echo    Aplicatia porneste...
echo.
echo    Local:   http://localhost:5050
echo.
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
    echo    Retea:   http://%%a:5050
)
echo.
echo    Pentru oprire: inchideti aceasta fereastra
echo ============================================
echo.

start "" http://localhost:5050
python app.py
pause

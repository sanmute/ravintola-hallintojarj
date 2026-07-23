@echo off
chcp 65001 > nul
echo ============================================
echo   Ruokalistasuunnittelija
echo ============================================
cd /d "%~dp0"
python -c "import flask, openpyxl, requests, bs4, docx" 2>nul || (
    echo Asennetaan/paivitetaan riippuvuudet...
    pip install -r requirements.txt
)
echo.
echo Kaynnistetaan... Avaa selaimessa: http://localhost:5001
echo Sulje painamalla Ctrl+C
echo.
python app.py
pause

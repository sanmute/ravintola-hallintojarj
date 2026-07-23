@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo Lisataan aloitusreseptit (talvi)...
python seed_recipes.py talvi
echo.
echo Valmis! Kaynnista sovellus tiedostolla KAYNNISTA.bat
pause

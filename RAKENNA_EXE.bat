@echo off
chcp 65001 >nul
echo ============================================================
echo   Ruokalistasuunnittelija - EXE-paketointi
echo ============================================================
echo.

echo [1/3] Asennetaan paketointityokalut...
python -m pip install pyinstaller pywebview --quiet
if errorlevel 1 (
    echo VIRHE: pip-asennus epaonnistui.
    pause
    exit /b 1
)

echo [2/3] Rakennetaan sovellus (kestaa muutaman minuutin)...
python -m PyInstaller ruokalistasuunnittelija.spec --noconfirm
if errorlevel 1 (
    echo VIRHE: paketointi epaonnistui. Katso virheilmoitus yllaolta.
    pause
    exit /b 1
)

echo [3/3] Valmis!
echo.
echo Sovellus loytyy kansiosta:
echo   dist\Ruokalistasuunnittelija\Ruokalistasuunnittelija.exe
echo.
echo Kopioi KOKO dist\Ruokalistasuunnittelija-kansio kohdekoneelle
echo ja tee tyopoydalle pikakuvake exe-tiedostoon.
echo.
pause
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

echo [3/4] Valmis!
echo.
echo Sovellus loytyy tiedostosta (YKSI tiedosto, onefile-tila):
echo   dist\Ruokalistasuunnittelija.exe
echo.
echo Voit kopioida taman yhden exe-tiedoston suoraan kohdekoneelle.
echo Huom: kaynnistys kestaa muutaman sekunnin, koska sovellus puretaan
echo Temp-kansioon joka kaynnistyskerralla (onefile-tilan kompromissi).
echo.

echo [4/4] Luodaan tyopoydalle pikakuvake...
rem Onefile-tilassa tama ei ole enaa valttamatonta yhden .exe:n irtoamisen
rem kannalta (tiedosto on jo itsenaisesti kopioitavissa), mutta pikakuvake
rem on silti mukavampi kayttaa kuin exe-tiedosto suoraan.
set "EXE_DIR=%~dp0dist"
set "EXE_PATH=%EXE_DIR%\Ruokalistasuunnittelija.exe"
if exist "%EXE_PATH%" (
    powershell -NoProfile -Command ^
        "$s = (New-Object -ComObject WScript.Shell).CreateShortcut(\"$env:USERPROFILE\Desktop\Ruokalistasuunnittelija.lnk\"); $s.TargetPath = '%EXE_PATH%'; $s.WorkingDirectory = '%EXE_DIR%'; $s.IconLocation = '%EXE_PATH%'; $s.Save()"
    echo Pikakuvake luotu tyopoydalle: Ruokalistasuunnittelija.lnk
) else (
    echo VAROITUS: exe-tiedostoa ei loytynyt, pikakuvaketta ei luotu.
)
echo.
pause
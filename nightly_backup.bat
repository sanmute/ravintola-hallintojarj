@echo off
REM Windows Task Scheduler -kaynnistin: yollinen tietokannan varmuuskopiointi.
REM Kutsuu nightly_backup.py:ta ja kirjoittaa tuloksen varmuuskopiot\backup_log.txt-tiedostoon.
REM Tama tiedosto EI ole sama kuin sovelluksen sisainen backup.py-moduuli
REM (joka tarjoaa /api/backup-reitit) - tama on itsenainen, sovelluksesta riippumaton skripti.

setlocal

set SCRIPT_DIR=%~dp0
set LOG_FILE=%SCRIPT_DIR%varmuuskopiot\backup_log.txt

if not exist "%SCRIPT_DIR%varmuuskopiot" mkdir "%SCRIPT_DIR%varmuuskopiot"

echo [%date% %time%] Ajetaan varmuuskopiointia... >> "%LOG_FILE%"

REM Yrita ensin PATH:issa olevaa pythonia; Task Schedulerin ajoymparistossa
REM PATH voi kuitenkin poiketa interaktiivisesta kuoresta, joten varalla
REM kaytetaan tunnettua asennuspolkua.
python "%SCRIPT_DIR%nightly_backup.py" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    "C:\Users\santeri.mutanen\AppData\Local\Python\pythoncore-3.14-64\python.exe" "%SCRIPT_DIR%nightly_backup.py" >> "%LOG_FILE%" 2>&1
)

echo [%date% %time%] Valmis. >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

endlocal

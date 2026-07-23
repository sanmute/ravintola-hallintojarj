# setup_nightly_backup_task.ps1
# Rekisteroi Windowsin Task Scheduleriin ajastetun tehtavan, joka ajaa
# nightly_backup.bat joka ilta klo 22:00 -- riippumatta siita onko
# Ruokalistasuunnittelija-sovellus itse kaynnissa.
#
# Kaytto (kertaalleen):
#   1. Avaa PowerShell TAVALLISENA kayttajana (admin-oikeuksia ei tarvita
#      talle oletusasetukselle: tehtava ajetaan vain kirjautuneena ollessa).
#   2. Aja tasta kansiosta:  .\setup_nightly_backup_task.ps1
#
# Jos haluat tehtavan ajautuvan myos silloin kun kukaan ei ole kirjautuneena
# sisaan, tarvitset admin-oikeudet ja -User/-Password- tai -RunLevel Highest
# -parametrit Register-ScheduledTask-kutsuun -- ei toteutettu tassa,
# koska se vaatisi salasanan tallentamista.

$ErrorActionPreference = 'Stop'

$taskName = 'Ruokalistasuunnittelija Backup'
$scriptDir = $PSScriptRoot
$batFile = Join-Path $scriptDir 'nightly_backup.bat'

if (-not (Test-Path $batFile)) {
    Write-Error "nightly_backup.bat ei loytynyt polusta: $batFile"
    exit 1
}

$trigger = New-ScheduledTaskTrigger -Daily -At 22:00
$action = New-ScheduledTaskAction -Execute $batFile -WorkingDirectory $scriptDir
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask -TaskName $taskName `
    -Trigger $trigger `
    -Action $action `
    -Settings $settings `
    -Description 'Automaattinen yollinen tietokannan varmuuskopiointi (Ruokalistasuunnittelija).' `
    -Force | Out-Null

Write-Host "Tehtava '$taskName' luotu."
Write-Host 'Ajastettu joka ilta klo 22:00.'
Write-Host "Loki: $scriptDir\varmuuskopiot\backup_log.txt"
Write-Host ''
Write-Host 'Testaa heti (valinnainen):'
Write-Host "  Start-ScheduledTask -TaskName '$taskName'"

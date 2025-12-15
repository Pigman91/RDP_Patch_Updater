#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Nainstaluje naplanovanu ulohu pro automatickou aktualizaci RDP Wrapper

.DESCRIPTION
    Vytvori naplanovanu ulohu, ktera se spusti pri startu systemu

.PARAMETER Uninstall
    Odstrani naplanovanu ulohu

.PARAMETER AutoRestart
    Pokud je zadano, naplanovan uloha bude automaticky restartovat PC po aktualizaci

.PARAMETER UserName
    Jmeno uzivatele pod kterym ma uloha bezet

.PARAMETER Password
    Heslo uzivatele (pokud neni zadano, bude vyzadano)

.EXAMPLE
    .\Install-ScheduledTask.ps1 -AutoRestart -UserName "Administrator"
    .\Install-ScheduledTask.ps1 -AutoRestart -UserName "DOMAIN\Administrator"
    .\Install-ScheduledTask.ps1 -Uninstall
#>

param(
    [switch]$Uninstall,
    [switch]$AutoRestart,
    [string]$UserName,
    [string]$Password
)

$TaskName = "RDP Wrapper Auto Update"

# Pevna cesta k instalacnimu adresari
$InstallDir = "C:\Program Files\RDP Wrapper\AutoUpdate"
$UpdateScript = Join-Path $InstallDir "Update-RDPWrap.ps1"

if ($Uninstall) {
    Write-Host "Odstranuji naplanovanu ulohu '$TaskName'..." -ForegroundColor Yellow
    schtasks /Delete /TN "$TaskName" /F 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Naplanovan uloha odstranena" -ForegroundColor Green
    } else {
        Write-Host "Naplanovan uloha neexistuje nebo se nepodarilo odstranit" -ForegroundColor Yellow
    }
    exit 0
}

# Kontrola existence Update skriptu
if (-not (Test-Path $UpdateScript)) {
    Write-Host "CHYBA: Update-RDPWrap.ps1 nenalezen v: $UpdateScript" -ForegroundColor Red
    exit 1
}

# Odstranit existujici ulohu pokud existuje
schtasks /Delete /TN "$TaskName" /F 2>$null

# Pripravit prikaz pro PowerShell (cela cesta v uvozovkach)
$Command = if ($AutoRestart) {
    "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File \`"$UpdateScript\`" -AutoRestart"
} else {
    "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File \`"$UpdateScript\`""
}

# Urcit uzivatele
if (-not $UserName) {
    $UserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
}

# Zeptat se na heslo pokud nebylo zadano
if (-not $Password) {
    Write-Host ""
    Write-Host "Pro spusteni ulohy pri startu systemu je potreba heslo uzivatele." -ForegroundColor Yellow
    Write-Host "Uzivatel: $UserName" -ForegroundColor Cyan
    Write-Host ""
    $SecurePassword = Read-Host -Prompt "Zadej heslo" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
}

# Vytvorit ulohu pomoci schtasks
Write-Host ""
Write-Host "Vytvarim naplanovanu ulohu..." -ForegroundColor Yellow

$Result = schtasks /Create /TN "$TaskName" /TR $Command /SC ONSTART /RU "$UserName" /RP "$Password" /RL HIGHEST /F 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "Naplanovan uloha uspesne vytvorena!" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Nazev: $TaskName"
    Write-Host "Spousti se: Pri startu systemu"
    Write-Host "Uzivatel: $UserName"
    Write-Host "Skript: $UpdateScript"

    if ($AutoRestart) {
        Write-Host ""
        Write-Host "POZOR: AutoRestart je AKTIVNI!" -ForegroundColor Yellow
        Write-Host "Po nalezeni novych offsetu se PC automaticky restartuje." -ForegroundColor Yellow
    }
    else {
        Write-Host ""
        Write-Host "AutoRestart je vypnuty." -ForegroundColor Cyan
        Write-Host "Nove offsety budou pridany, ale PC se nerestartuje automaticky."
    }

    Write-Host ""
    Write-Host "Pro odinstalaci spustte: .\Install-ScheduledTask.ps1 -Uninstall" -ForegroundColor Gray
}
else {
    Write-Host "CHYBA pri vytvareni napl novane ulohy:" -ForegroundColor Red
    Write-Host $Result -ForegroundColor Red
    exit 1
}

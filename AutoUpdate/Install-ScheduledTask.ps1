#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs a scheduled task for automatic RDP Wrapper updates

.DESCRIPTION
    Creates a scheduled task that runs at system startup

.PARAMETER Uninstall
    Removes the scheduled task

.PARAMETER AutoRestart
    If specified, the scheduled task will automatically restart PC after update

.PARAMETER UserName
    Username under which the task will run

.PARAMETER Password
    User password (if not provided, will be requested)

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

# Fixed path to installation directory
$InstallDir = "C:\Program Files\RDP Wrapper\AutoUpdate"
$UpdateScript = Join-Path $InstallDir "Update-RDPWrap.ps1"

if ($Uninstall) {
    Write-Host "Removing scheduled task '$TaskName'..." -ForegroundColor Yellow
    schtasks /Delete /TN "$TaskName" /F 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Scheduled task removed" -ForegroundColor Green
    } else {
        Write-Host "Scheduled task does not exist or could not be removed" -ForegroundColor Yellow
    }
    exit 0
}

# Check if Update script exists
if (-not (Test-Path $UpdateScript)) {
    Write-Host "ERROR: Update-RDPWrap.ps1 not found at: $UpdateScript" -ForegroundColor Red
    exit 1
}

# Remove existing task if exists
schtasks /Delete /TN "$TaskName" /F 2>$null

# Prepare PowerShell command (full path in quotes)
$Command = if ($AutoRestart) {
    "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File \`"$UpdateScript\`" -AutoRestart"
} else {
    "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File \`"$UpdateScript\`""
}

# Determine user
if (-not $UserName) {
    $UserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
}

# Request password if not provided
if (-not $Password) {
    Write-Host ""
    Write-Host "User password is required to run the task at system startup." -ForegroundColor Yellow
    Write-Host "User: $UserName" -ForegroundColor Cyan
    Write-Host ""
    $SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
}

# Create task using schtasks
Write-Host ""
Write-Host "Creating scheduled task..." -ForegroundColor Yellow

$Result = schtasks /Create /TN "$TaskName" /TR $Command /SC ONSTART /RU "$UserName" /RP "$Password" /RL HIGHEST /F 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "Scheduled task created successfully!" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Name: $TaskName"
    Write-Host "Trigger: At system startup"
    Write-Host "User: $UserName"
    Write-Host "Script: $UpdateScript"

    if ($AutoRestart) {
        Write-Host ""
        Write-Host "WARNING: AutoRestart is ENABLED!" -ForegroundColor Yellow
        Write-Host "PC will automatically restart after finding new offsets." -ForegroundColor Yellow
    }
    else {
        Write-Host ""
        Write-Host "AutoRestart is disabled." -ForegroundColor Cyan
        Write-Host "New offsets will be added, but PC will not restart automatically."
    }

    Write-Host ""
    Write-Host "To uninstall run: .\Install-ScheduledTask.ps1 -Uninstall" -ForegroundColor Gray
}
else {
    Write-Host "ERROR creating scheduled task:" -ForegroundColor Red
    Write-Host $Result -ForegroundColor Red
    exit 1
}

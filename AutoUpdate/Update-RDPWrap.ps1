#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Automatically updates rdpwrap.ini with new offsets for current termsrv.dll version

.DESCRIPTION
    Script runs RDPWrapOffsetFinder.exe, gets new offsets and if they are not yet
    in rdpwrap.ini, adds them. Optionally restarts the computer.

.PARAMETER AutoRestart
    If specified, restarts the computer after successful update

.PARAMETER Force
    Adds offsets even if they already exist (overwrites)

.EXAMPLE
    .\Update-RDPWrap.ps1
    .\Update-RDPWrap.ps1 -AutoRestart
#>

param(
    [switch]$AutoRestart,
    [switch]$Force
)

# Configuration
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# Fallback to fixed path if $ScriptDir doesn't evaluate correctly
if ([string]::IsNullOrEmpty($ScriptDir) -or -not (Test-Path $ScriptDir)) {
    $ScriptDir = "C:\Program Files\RDP Wrapper\AutoUpdate"
}
$OffsetFinderPath = Join-Path $ScriptDir "OffsetFinder\RDPWrapOffsetFinder.exe"
$TermsrvPath = "C:\Windows\System32\termsrv.dll"
$RdpWrapIniPath = "C:\Program Files\RDP Wrapper\rdpwrap.ini"
$LogFile = Join-Path $ScriptDir "Update-RDPWrap.log"

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage -Encoding UTF8

    switch ($Level) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "WARN"  { Write-Host $LogMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $LogMessage -ForegroundColor Green }
        default { Write-Host $LogMessage }
    }
}

# Check prerequisites
function Test-Prerequisites {
    Write-Log "Checking prerequisites..."

    if (-not (Test-Path $OffsetFinderPath)) {
        Write-Log "RDPWrapOffsetFinder.exe not found: $OffsetFinderPath" "ERROR"
        return $false
    }

    if (-not (Test-Path $TermsrvPath)) {
        Write-Log "termsrv.dll not found: $TermsrvPath" "ERROR"
        return $false
    }

    if (-not (Test-Path $RdpWrapIniPath)) {
        Write-Log "rdpwrap.ini not found: $RdpWrapIniPath" "ERROR"
        return $false
    }

    Write-Log "All prerequisites met"
    return $true
}

# Get new offsets
function Get-NewOffsets {
    Write-Log "Running RDPWrapOffsetFinder.exe..."

    $OffsetFinderDir = Split-Path -Parent $OffsetFinderPath
    $TempOutputFile = Join-Path $ScriptDir "offsetfinder_output.tmp"

    try {
        # Use Start-Process with redirection to file - more reliable than & operator
        Write-Log "Trying version with symbols..."
        $ExePath = Join-Path $OffsetFinderDir "RDPWrapOffsetFinder.exe"

        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
        $ProcessInfo.FileName = $ExePath
        $ProcessInfo.Arguments = "`"$TermsrvPath`""
        $ProcessInfo.WorkingDirectory = $OffsetFinderDir
        $ProcessInfo.RedirectStandardOutput = $true
        $ProcessInfo.RedirectStandardError = $true
        $ProcessInfo.UseShellExecute = $false
        $ProcessInfo.CreateNoWindow = $true

        $Process = New-Object System.Diagnostics.Process
        $Process.StartInfo = $ProcessInfo
        $Process.Start() | Out-Null

        $OutputText = $Process.StandardOutput.ReadToEnd()
        $ErrorText = $Process.StandardError.ReadToEnd()
        $Process.WaitForExit()

        $ExitCode = $Process.ExitCode

        # If there's an error in output, try nosymbol version
        if ($OutputText -match "ERROR:" -or $ErrorText -match "ERROR:" -or [string]::IsNullOrWhiteSpace($OutputText)) {
            Write-Log "Symbol version failed (or empty output), trying nosymbol version..." "WARN"

            $NoSymbolPath = Join-Path $OffsetFinderDir "RDPWrapOffsetFinder_nosymbol.exe"
            if (Test-Path $NoSymbolPath) {
                $ProcessInfo.FileName = $NoSymbolPath

                $Process = New-Object System.Diagnostics.Process
                $Process.StartInfo = $ProcessInfo
                $Process.Start() | Out-Null

                $OutputText = $Process.StandardOutput.ReadToEnd()
                $ErrorText = $Process.StandardError.ReadToEnd()
                $Process.WaitForExit()

                $ExitCode = $Process.ExitCode
            }
        }

        if ($ExitCode -ne 0) {
            Write-Log "RDPWrapOffsetFinder failed with code: $ExitCode" "ERROR"
            return $null
        }

        # Final check for ERROR
        if ($OutputText -match "ERROR:") {
            Write-Log "OffsetFinder could not find all offsets" "ERROR"
            Write-Log "Output: $OutputText" "ERROR"
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($OutputText)) {
            Write-Log "OffsetFinder returned empty output" "ERROR"
            return $null
        }

        Write-Log "Output retrieved successfully"
        return $OutputText
    }
    catch {
        Write-Log "Error running RDPWrapOffsetFinder: $_" "ERROR"
        return $null
    }
}

# Extract version from output
function Get-VersionFromOutput {
    param([string]$Output)

    # Looking for pattern [10.0.xxxxx.xxxx]
    if ($Output -match '\[(\d+\.\d+\.\d+\.\d+)\]') {
        return $Matches[1]
    }
    return $null
}

# Check if version exists in INI
function Test-VersionExists {
    param([string]$Version, [string]$IniContent)

    $Pattern = "\[$([regex]::Escape($Version))\]"
    return $IniContent -match $Pattern
}

# Create backup
function Backup-IniFile {
    $BackupDir = Join-Path $ScriptDir "Backups"
    if (-not (Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
    }

    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $BackupPath = Join-Path $BackupDir "rdpwrap_$Timestamp.ini"

    Copy-Item -Path $RdpWrapIniPath -Destination $BackupPath -Force
    Write-Log "Backup created: $BackupPath"

    return $BackupPath
}

# Add new offsets to INI
function Add-OffsetsToIni {
    param([string]$NewOffsets, [string]$BackupPath)

    # Use backup as source (it already exists and is not locked)
    $CurrentContent = Get-Content -Path $BackupPath -Raw -Encoding UTF8

    # Remove any trailing empty lines and add new offsets
    $CurrentContent = $CurrentContent.TrimEnd()
    $NewContent = $CurrentContent + "`r`n`r`n" + $NewOffsets.Trim() + "`r`n"

    # Create temporary file with new content
    $TempFile = Join-Path $ScriptDir "rdpwrap_new.ini"
    Set-Content -Path $TempFile -Value $NewContent -Encoding UTF8 -NoNewline
    Write-Log "Temporary file created: $TempFile"

    # Copy temporary file over original (this works even when file is in use)
    Copy-Item -Path $TempFile -Destination $RdpWrapIniPath -Force
    Write-Log "New offsets added to rdpwrap.ini" "SUCCESS"

    # Delete temporary file
    Remove-Item -Path $TempFile -Force
}

# Main logic
function Main {
    Write-Log "========== Starting Update-RDPWrap =========="

    # Check prerequisites
    if (-not (Test-Prerequisites)) {
        return 1
    }

    # Get new offsets
    $NewOffsets = Get-NewOffsets
    if (-not $NewOffsets) {
        Write-Log "Failed to get new offsets" "ERROR"
        return 1
    }

    # Extract version
    $Version = Get-VersionFromOutput -Output $NewOffsets
    if (-not $Version) {
        Write-Log "Failed to extract version from output" "ERROR"
        Write-Log "Output: $NewOffsets" "ERROR"
        return 1
    }

    Write-Log "Found version: $Version"

    # Load current INI and check if version exists
    $CurrentIni = Get-Content -Path $RdpWrapIniPath -Raw -Encoding UTF8

    if ((Test-VersionExists -Version $Version -IniContent $CurrentIni) -and -not $Force) {
        Write-Log "Version $Version already exists in rdpwrap.ini - no action needed" "SUCCESS"
        return 0
    }

    if ($Force -and (Test-VersionExists -Version $Version -IniContent $CurrentIni)) {
        Write-Log "Version $Version already exists, but Force is active - continuing" "WARN"
    }

    # Create backup
    $BackupPath = Backup-IniFile

    # Add new offsets
    try {
        Add-OffsetsToIni -NewOffsets $NewOffsets -BackupPath $BackupPath
    }
    catch {
        Write-Log "Error writing to rdpwrap.ini: $_" "ERROR"
        Write-Log "Restoring from backup..." "WARN"
        Copy-Item -Path $BackupPath -Destination $RdpWrapIniPath -Force
        return 1
    }

    Write-Log "Update completed successfully" "SUCCESS"

    # Restart if requested
    if ($AutoRestart) {
        Write-Log "AutoRestart is active - restarting computer in 30 seconds..." "WARN"
        Write-Log "To cancel run: shutdown /a"
        shutdown /r /t 30 /c "RDP Wrapper updated - system restart"
    }
    else {
        Write-Log "Computer restart is required to activate new offsets" "WARN"
    }

    return 0
}

# Run
$ExitCode = Main
exit $ExitCode

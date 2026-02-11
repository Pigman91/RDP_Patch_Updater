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
    [switch]$Force,
    [switch]$SkipSelfUpdate
)

# Script version
$ScriptVersion = "2.5.0"

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

# Self-update configuration
$SelfUpdateUrl = "https://pigman91.github.io/RDP_Patch_Updater/AutoUpdate/Update-RDPWrap.ps1"
$ScriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = Join-Path $ScriptDir "Update-RDPWrap.ps1"
}

# Retry configuration - progressive delays in seconds
# Attempt 1: immediate, then 2min, 5min, 15min, 30min, 60min (total ~112 min)
$RetryDelays = @(120, 300, 900, 1800, 3600)
$NetworkCheckTimeout = 300  # Max 5 min wait for network at startup

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

# Test if network (symbol server) is reachable
function Test-NetworkReady {
    $Hosts = @("msdl.microsoft.com", "microsoft.com")
    foreach ($H in $Hosts) {
        try {
            $Tcp = New-Object System.Net.Sockets.TcpClient
            $AsyncResult = $Tcp.BeginConnect($H, 443, $null, $null)
            $Wait = $AsyncResult.AsyncWaitHandle.WaitOne(5000, $false)
            if ($Wait -and $Tcp.Connected) {
                $Tcp.Close()
                return $true
            }
            $Tcp.Close()
        }
        catch {}
    }
    return $false
}

# Wait for network connectivity before proceeding
function Wait-ForNetwork {
    Write-Log "Checking network connectivity..."

    if (Test-NetworkReady) {
        Write-Log "Network is available"
        return $true
    }

    $Elapsed = 0
    $CheckInterval = 15
    while ($Elapsed -lt $NetworkCheckTimeout) {
        Write-Log "Network not ready, waiting ${CheckInterval}s... (${Elapsed}s/${NetworkCheckTimeout}s)" "WARN"
        Start-Sleep -Seconds $CheckInterval
        $Elapsed += $CheckInterval

        if (Test-NetworkReady) {
            Write-Log "Network is available (after ${Elapsed}s)"
            return $true
        }
    }

    Write-Log "Network not available after ${NetworkCheckTimeout}s" "ERROR"
    return $false
}

# Extract PDB GUID from DLL PE headers (needed to download PDB from Microsoft Symbol Server)
function Get-PdbInfo {
    param([string]$DllPath)

    try {
        $stream = [System.IO.File]::OpenRead($DllPath)
        $reader = New-Object System.IO.BinaryReader($stream)

        # DOS header - get PE offset
        $stream.Position = 60
        $peOffset = $reader.ReadUInt32()

        # COFF header (skip PE signature)
        $stream.Position = $peOffset + 4
        $reader.ReadUInt16() | Out-Null  # machine
        $numberOfSections = $reader.ReadUInt16()
        $reader.ReadBytes(12) | Out-Null  # skip timeDateStamp, pointerToSymbolTable, numberOfSymbols
        $sizeOfOptionalHeader = $reader.ReadUInt16()
        $reader.ReadUInt16() | Out-Null  # characteristics

        # Optional header
        $optionalHeaderOffset = $stream.Position
        $magic = $reader.ReadUInt16()
        $is64Bit = ($magic -eq 0x20B)

        # Debug directory is 7th data directory entry (index 6)
        $dataDirOffset = $optionalHeaderOffset + $(if ($is64Bit) { 112 } else { 96 }) + (6 * 8)
        $stream.Position = $dataDirOffset
        $debugDirRva = $reader.ReadUInt32()
        $debugDirSize = $reader.ReadUInt32()

        if ($debugDirRva -eq 0) {
            $reader.Close(); $stream.Close()
            return $null
        }

        # Find section containing debug directory
        $sectionHeaderOffset = $optionalHeaderOffset + $sizeOfOptionalHeader
        $debugDirFileOffset = 0

        for ($i = 0; $i -lt $numberOfSections; $i++) {
            $stream.Position = $sectionHeaderOffset + ($i * 40) + 8
            $virtualSize = $reader.ReadUInt32()
            $sectionRva = $reader.ReadUInt32()
            $reader.ReadUInt32() | Out-Null  # sizeOfRawData
            $pointerToRawData = $reader.ReadUInt32()

            if ($debugDirRva -ge $sectionRva -and $debugDirRva -lt ($sectionRva + $virtualSize)) {
                $debugDirFileOffset = $pointerToRawData + ($debugDirRva - $sectionRva)
                break
            }
        }

        if ($debugDirFileOffset -eq 0) {
            $reader.Close(); $stream.Close()
            return $null
        }

        # Read debug directory entries looking for CodeView (type 2)
        $stream.Position = $debugDirFileOffset
        $numEntries = [int]($debugDirSize / 28)

        for ($i = 0; $i -lt $numEntries; $i++) {
            $reader.ReadBytes(12) | Out-Null  # characteristics, timeDateStamp, version
            $type = $reader.ReadUInt32()
            $reader.ReadUInt32() | Out-Null  # sizeOfData
            $reader.ReadUInt32() | Out-Null  # addressOfRawData
            $ptrToRawData = $reader.ReadUInt32()

            if ($type -eq 2) {  # IMAGE_DEBUG_TYPE_CODEVIEW
                $savedPos = $stream.Position
                $stream.Position = $ptrToRawData
                $cvSig = $reader.ReadUInt32()

                if ($cvSig -eq 0x53445352) {  # "RSDS"
                    $guidBytes = $reader.ReadBytes(16)
                    $guid = New-Object System.Guid(,$guidBytes)
                    $age = $reader.ReadUInt32()

                    $nameBytes = New-Object System.Collections.ArrayList
                    while ($true) {
                        $b = $reader.ReadByte()
                        if ($b -eq 0) { break }
                        [void]$nameBytes.Add($b)
                    }
                    $pdbName = [System.IO.Path]::GetFileName(
                        [System.Text.Encoding]::ASCII.GetString($nameBytes.ToArray())
                    )
                    $symbolId = $guid.ToString("N").ToUpper() + $age.ToString()

                    $reader.Close(); $stream.Close()
                    return @{
                        PdbName  = $pdbName
                        SymbolId = $symbolId
                        Url      = "https://msdl.microsoft.com/download/symbols/$pdbName/$symbolId/$pdbName"
                    }
                }
                $stream.Position = $savedPos
            }
        }

        $reader.Close(); $stream.Close()
        return $null
    }
    catch {
        return $null
    }
}

# Pre-download PDB from Microsoft Symbol Server using PowerShell HTTP client
# (symsrv.dll inside OffsetFinder uses WinINet which fails in non-interactive Task Scheduler sessions)
function Ensure-PdbCached {
    $OffsetFinderDir = Split-Path -Parent $OffsetFinderPath
    $SymDir = Join-Path $OffsetFinderDir "sym"

    $PdbInfo = Get-PdbInfo -DllPath $TermsrvPath
    if (-not $PdbInfo) {
        Write-Log "Could not extract PDB info from termsrv.dll (will rely on OffsetFinder)" "WARN"
        return $false
    }

    Write-Log "PDB: $($PdbInfo.PdbName) [$($PdbInfo.SymbolId)]"

    # Check if PDB is already cached
    $PdbCachePath = Join-Path $SymDir "$($PdbInfo.PdbName)\$($PdbInfo.SymbolId)\$($PdbInfo.PdbName)"
    if (Test-Path $PdbCachePath) {
        $PdbSize = (Get-Item $PdbCachePath).Length
        if ($PdbSize -gt 0) {
            Write-Log "PDB already cached ($([math]::Round($PdbSize / 1MB, 2)) MB)"
            return $true
        }
        # Zero-byte PDB = corrupted, remove it
        Remove-Item -Path $PdbCachePath -Force -ErrorAction SilentlyContinue
    }

    # Download PDB via PowerShell (works in non-interactive sessions unlike symsrv.dll)
    Write-Log "Downloading PDB from Microsoft Symbol Server..."
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $PdbDir = Split-Path -Parent $PdbCachePath
        if (-not (Test-Path $PdbDir)) {
            New-Item -ItemType Directory -Path $PdbDir -Force | Out-Null
        }

        $SW = [System.Diagnostics.Stopwatch]::StartNew()
        Invoke-WebRequest -Uri $PdbInfo.Url -OutFile $PdbCachePath -UseBasicParsing -TimeoutSec 120
        $SW.Stop()

        $DownloadedSize = (Get-Item $PdbCachePath).Length
        if ($DownloadedSize -gt 0) {
            Write-Log "PDB downloaded: $([math]::Round($DownloadedSize / 1MB, 2)) MB in $([math]::Round($SW.Elapsed.TotalSeconds, 1))s" "SUCCESS"
            return $true
        }
        else {
            Write-Log "PDB download returned empty file" "WARN"
            Remove-Item -Path $PdbCachePath -Force -ErrorAction SilentlyContinue
            return $false
        }
    }
    catch {
        Write-Log "PDB download failed: $_ (will rely on OffsetFinder)" "WARN"
        Remove-Item -Path $PdbCachePath -Force -ErrorAction SilentlyContinue
        return $false
    }
}

# Self-update: check GitHub for newer script version
function Update-Self {
    Write-Log "Checking for script updates..."

    $TempScript = Join-Path $ScriptDir "Update-RDPWrap_update.ps1"

    try {
        Invoke-WebRequest -Uri $SelfUpdateUrl -OutFile $TempScript -UseBasicParsing -TimeoutSec 30

        $LocalHash = (Get-FileHash -Path $ScriptPath -Algorithm SHA256).Hash
        $RemoteHash = (Get-FileHash -Path $TempScript -Algorithm SHA256).Hash

        if ($LocalHash -ne $RemoteHash) {
            Write-Log "New script version found, updating..."
            Copy-Item -Path $TempScript -Destination $ScriptPath -Force
            Write-Log "Script updated successfully, restarting with new version..." "SUCCESS"

            # Relaunch with same arguments + SkipSelfUpdate to prevent loop
            $RelaunchArgs = "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$ScriptPath`" -SkipSelfUpdate"
            if ($AutoRestart) { $RelaunchArgs += " -AutoRestart" }
            if ($Force) { $RelaunchArgs += " -Force" }

            Start-Process "powershell.exe" -ArgumentList $RelaunchArgs
            return $true
        }

        Write-Log "Script is up to date"
        return $false
    }
    catch {
        Write-Log "Self-update check failed: $_ (continuing with current version)" "WARN"
        return $false
    }
    finally {
        Remove-Item -Path $TempScript -Force -ErrorAction SilentlyContinue
    }
}

# Store partial output from last OffsetFinder run (for version extraction on failure)
$script:LastOffsetFinderOutput = ""

# Run OffsetFinder once (internal function)
function Invoke-OffsetFinder {
    $OffsetFinderDir = Split-Path -Parent $OffsetFinderPath

    try {
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

        # Save partial output for version extraction even on failure
        if (-not [string]::IsNullOrWhiteSpace($OutputText)) {
            $script:LastOffsetFinderOutput = $OutputText
        }

        # If there's an error in output, try nosymbol version
        if ($OutputText -match "ERROR:" -or $ErrorText -match "ERROR:" -or [string]::IsNullOrWhiteSpace($OutputText)) {
            Write-Log "Symbol version failed (exit=$ExitCode)" "WARN"
            if (-not [string]::IsNullOrWhiteSpace($OutputText)) { Write-Log "Symbol stdout: $($OutputText.Trim())" "WARN" }
            if (-not [string]::IsNullOrWhiteSpace($ErrorText)) { Write-Log "Symbol stderr: $($ErrorText.Trim())" "WARN" }

            Write-Log "Trying nosymbol version..."
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
            if (-not [string]::IsNullOrWhiteSpace($OutputText)) { Write-Log "stdout: $($OutputText.Trim())" "ERROR" }
            if (-not [string]::IsNullOrWhiteSpace($ErrorText)) { Write-Log "stderr: $($ErrorText.Trim())" "ERROR" }
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

# Retry OffsetFinder with progressive delays (first attempt already done in Main)
function Get-NewOffsetsRetry {
    $TotalAttempts = $RetryDelays.Count

    for ($i = 0; $i -lt $RetryDelays.Count; $i++) {
        $Delay = $RetryDelays[$i]
        $DelayMin = [math]::Round($Delay / 60, 1)
        $AttemptNum = $i + 2

        Write-Log "Attempt $AttemptNum of $($TotalAttempts + 1) - waiting ${DelayMin} minutes before retry..." "WARN"
        Start-Sleep -Seconds $Delay

        Ensure-PdbCached
        $Result = Invoke-OffsetFinder
        if ($Result) {
            Write-Log "Succeeded on attempt $AttemptNum of $($TotalAttempts + 1)"
            return $Result
        }
    }

    Write-Log "All $($TotalAttempts + 1) attempts failed" "ERROR"
    return $null
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
    Write-Log "========== Starting Update-RDPWrap v$ScriptVersion =========="

    # Self-update check (skip if we just updated to prevent loop)
    if (-not $SkipSelfUpdate) {
        $WasUpdated = Update-Self
        if ($WasUpdated) {
            Write-Log "Exiting - new version has been launched"
            return 0
        }
    }

    # Check prerequisites
    if (-not (Test-Prerequisites)) {
        return 1
    }

    # Force re-check after script upgrade
    $LastHashFile = Join-Path $ScriptDir ".last_termsrv_hash"
    $LastVersionFile = Join-Path $ScriptDir ".last_script_version"
    $ScriptUpgraded = $false

    $ForceRecheck = $false
    if (Test-Path $LastVersionFile) {
        $LastVersion = (Get-Content $LastVersionFile -Raw).Trim()
        if ($LastVersion -ne $ScriptVersion) {
            Write-Log "Script upgraded ($LastVersion -> $ScriptVersion) - forcing full check"
            $ForceRecheck = $true
        }
    }
    else {
        Write-Log "No version marker found - forcing full check"
        $ForceRecheck = $true
    }

    if ($ForceRecheck) {
        Remove-Item -Path $LastHashFile -Force -ErrorAction SilentlyContinue
    }
    Set-Content -Path $LastVersionFile -Value $ScriptVersion

    # Early check - skip if termsrv.dll hasn't changed since last successful run
    if (-not $Force) {
        try {
            $TermsrvHash = (Get-FileHash $TermsrvPath -Algorithm SHA256).Hash
            Write-Log "Current termsrv.dll hash: $TermsrvHash"

            if (Test-Path $LastHashFile) {
                $LastHash = (Get-Content $LastHashFile -Raw).Trim()
                if ($LastHash -eq $TermsrvHash) {
                    Write-Log "termsrv.dll unchanged since last successful check - no action needed" "SUCCESS"
                    return 0
                }
                Write-Log "termsrv.dll has changed since last check - need to find offsets"
            }
            else {
                Write-Log "First run - need to verify offsets"
            }
        }
        catch {
            Write-Log "Could not pre-check hash: $_ (continuing with OffsetFinder)" "WARN"
        }
    }

    # Wait for network connectivity (important at system startup)
    if (-not (Wait-ForNetwork)) {
        Write-Log "Proceeding anyway, OffsetFinder may fail..." "WARN"
    }

    # Pre-download PDB from Microsoft Symbol Server
    # (symsrv.dll inside OffsetFinder uses WinINet which fails in non-interactive Task Scheduler sessions)
    $OffsetFinderDir = Split-Path -Parent $OffsetFinderPath
    $SymDir = Join-Path $OffsetFinderDir "sym"
    Ensure-PdbCached

    # Try OffsetFinder
    Write-Log "Running RDPWrapOffsetFinder.exe..."
    $NewOffsets = Invoke-OffsetFinder

    # If failed with "not found" errors, clear bad cache and retry with fresh PDB download
    if (-not $NewOffsets -and $script:LastOffsetFinderOutput -match "not found") {
        Write-Log "OffsetFinder failed - clearing cache and re-downloading PDB..." "WARN"

        # Remove entire sym/ directory
        if (Test-Path $SymDir) {
            Remove-Item -Path $SymDir -Recurse -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
            if (Test-Path $SymDir) {
                cmd /c "rmdir /s /q `"$SymDir`"" 2>$null
                Start-Sleep -Seconds 1
            }
        }

        # Remove stray pingme.txt (symsrv.dll negative cache marker)
        $PingmeRoot = Join-Path $OffsetFinderDir "pingme.txt"
        Remove-Item -Path $PingmeRoot -Force -ErrorAction SilentlyContinue

        # Re-download PDB via PowerShell HTTP and retry
        if (Ensure-PdbCached) {
            Write-Log "PDB re-downloaded, retrying OffsetFinder..."
            $NewOffsets = Invoke-OffsetFinder
            if ($NewOffsets) {
                Write-Log "Retry after PDB re-download succeeded" "SUCCESS"
            }
            else {
                Write-Log "Retry after PDB re-download still failed" "WARN"
            }
        }
    }

    # If first attempt failed, check alternatives before retrying
    if (-not $NewOffsets -and -not $Force) {
        $PartialVersion = Get-VersionFromOutput -Output $script:LastOffsetFinderOutput
        if ($PartialVersion) {
            # Check if version already exists in INI
            $CurrentIni = Get-Content -Path $RdpWrapIniPath -Raw -Encoding UTF8
            if (Test-VersionExists -Version $PartialVersion -IniContent $CurrentIni) {
                Write-Log "OffsetFinder failed but version $PartialVersion already exists in rdpwrap.ini - no action needed" "SUCCESS"
                try {
                    $SaveHash = (Get-FileHash $TermsrvPath -Algorithm SHA256).Hash
                    Set-Content -Path $LastHashFile -Value $SaveHash
                } catch {}
                return 0
            }

            Write-Log "Version $PartialVersion not in rdpwrap.ini - continuing with retries"
        }
    }

    # If still no offsets, retry OffsetFinder with progressive delays
    if (-not $NewOffsets) {
        $NewOffsets = Get-NewOffsetsRetry
        if (-not $NewOffsets) {
            Write-Log "Failed to get new offsets after all attempts" "ERROR"
            return 1
        }
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
        # Save hash so next run skips immediately
        try {
            $SaveHash = (Get-FileHash $TermsrvPath -Algorithm SHA256).Hash
            Set-Content -Path $LastHashFile -Value $SaveHash
        } catch {}
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

    # Save hash so next run skips immediately
    try {
        $SaveHash = (Get-FileHash $TermsrvPath -Algorithm SHA256).Hash
        Set-Content -Path $LastHashFile -Value $SaveHash
    } catch {}

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

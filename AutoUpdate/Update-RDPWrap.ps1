#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Automaticky aktualizuje rdpwrap.ini s novymi offsety pro aktualni verzi termsrv.dll

.DESCRIPTION
    Skript spusti RDPWrapOffsetFinder.exe, ziska nove offsety a pokud jeste nejsou
    v rdpwrap.ini, prida je. Volitelne restartuje pocitac.

.PARAMETER AutoRestart
    Pokud je zadano, po uspesne aktualizaci restartuje pocitac

.PARAMETER Force
    Prida offsety i kdyz uz existuji (prepise)

.EXAMPLE
    .\Update-RDPWrap.ps1
    .\Update-RDPWrap.ps1 -AutoRestart
#>

param(
    [switch]$AutoRestart,
    [switch]$Force
)

# Konfigurace
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# Fallback na pevnou cestu pokud se $ScriptDir nevyhodnoti spravne
if ([string]::IsNullOrEmpty($ScriptDir) -or -not (Test-Path $ScriptDir)) {
    $ScriptDir = "C:\Program Files\RDP Wrapper\AutoUpdate"
}
$OffsetFinderPath = Join-Path $ScriptDir "OffsetFinder\RDPWrapOffsetFinder.exe"
$TermsrvPath = "C:\Windows\System32\termsrv.dll"
$RdpWrapIniPath = "C:\Program Files\RDP Wrapper\rdpwrap.ini"
$LogFile = Join-Path $ScriptDir "Update-RDPWrap.log"

# Funkce pro logovani
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

# Kontrola predpokladu
function Test-Prerequisites {
    Write-Log "Kontroluji predpoklady..."

    if (-not (Test-Path $OffsetFinderPath)) {
        Write-Log "RDPWrapOffsetFinder.exe nenalezen: $OffsetFinderPath" "ERROR"
        return $false
    }

    if (-not (Test-Path $TermsrvPath)) {
        Write-Log "termsrv.dll nenalezen: $TermsrvPath" "ERROR"
        return $false
    }

    if (-not (Test-Path $RdpWrapIniPath)) {
        Write-Log "rdpwrap.ini nenalezen: $RdpWrapIniPath" "ERROR"
        return $false
    }

    Write-Log "Vsechny predpoklady splneny"
    return $true
}

# Ziskani novych offsetu
function Get-NewOffsets {
    Write-Log "Spoustim RDPWrapOffsetFinder.exe..."

    $OffsetFinderDir = Split-Path -Parent $OffsetFinderPath
    $TempOutputFile = Join-Path $ScriptDir "offsetfinder_output.tmp"

    try {
        # Pouzit Start-Process s presmerovanim do souboru - spolehlivejsi nez & operator
        Write-Log "Zkousim verzi se symboly..."
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

        # Pokud je chyba ve vystupu, zkusit nosymbol verzi
        if ($OutputText -match "ERROR:" -or $ErrorText -match "ERROR:" -or [string]::IsNullOrWhiteSpace($OutputText)) {
            Write-Log "Verze se symboly selhala (nebo prazdny vystup), zkousim nosymbol verzi..." "WARN"

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
            Write-Log "RDPWrapOffsetFinder selhal s kodem: $ExitCode" "ERROR"
            return $null
        }

        # Finalni kontrola na ERROR
        if ($OutputText -match "ERROR:") {
            Write-Log "OffsetFinder nedokazal najit vsechny offsety" "ERROR"
            Write-Log "Vystup: $OutputText" "ERROR"
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($OutputText)) {
            Write-Log "OffsetFinder vratil prazdny vystup" "ERROR"
            return $null
        }

        Write-Log "Vystup ziskan uspesne"
        return $OutputText
    }
    catch {
        Write-Log "Chyba pri spousteni RDPWrapOffsetFinder: $_" "ERROR"
        return $null
    }
}

# Extrakce verze z vystupu
function Get-VersionFromOutput {
    param([string]$Output)

    # Hledame vzor [10.0.xxxxx.xxxx]
    if ($Output -match '\[(\d+\.\d+\.\d+\.\d+)\]') {
        return $Matches[1]
    }
    return $null
}

# Kontrola zda verze existuje v INI
function Test-VersionExists {
    param([string]$Version, [string]$IniContent)

    $Pattern = "\[$([regex]::Escape($Version))\]"
    return $IniContent -match $Pattern
}

# Vytvoreni zalohy
function Backup-IniFile {
    $BackupDir = Join-Path $ScriptDir "Backups"
    if (-not (Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
    }

    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $BackupPath = Join-Path $BackupDir "rdpwrap_$Timestamp.ini"

    Copy-Item -Path $RdpWrapIniPath -Destination $BackupPath -Force
    Write-Log "Zaloha vytvorena: $BackupPath"

    return $BackupPath
}

# Pridani novych offsetu do INI
function Add-OffsetsToIni {
    param([string]$NewOffsets, [string]$BackupPath)

    # Pouzijeme zalohu jako zdroj (ta uz existuje a neni zamcena)
    $CurrentContent = Get-Content -Path $BackupPath -Raw -Encoding UTF8

    # Odstranit pripadne prazdne radky na konci a pridat nove offsety
    $CurrentContent = $CurrentContent.TrimEnd()
    $NewContent = $CurrentContent + "`r`n`r`n" + $NewOffsets.Trim() + "`r`n"

    # Vytvorit docasny soubor s novym obsahem
    $TempFile = Join-Path $ScriptDir "rdpwrap_new.ini"
    Set-Content -Path $TempFile -Value $NewContent -Encoding UTF8 -NoNewline
    Write-Log "Docasny soubor vytvoren: $TempFile"

    # Prekopirovat docasny soubor pres puvodni (toto funguje i kdyz je soubor pouzivan)
    Copy-Item -Path $TempFile -Destination $RdpWrapIniPath -Force
    Write-Log "Nove offsety pridany do rdpwrap.ini" "SUCCESS"

    # Smazat docasny soubor
    Remove-Item -Path $TempFile -Force
}

# Hlavni logika
function Main {
    Write-Log "========== Spoustim Update-RDPWrap =========="

    # Kontrola predpokladu
    if (-not (Test-Prerequisites)) {
        return 1
    }

    # Ziskat nove offsety
    $NewOffsets = Get-NewOffsets
    if (-not $NewOffsets) {
        Write-Log "Nepodarilo se ziskat nove offsety" "ERROR"
        return 1
    }

    # Extrahovat verzi
    $Version = Get-VersionFromOutput -Output $NewOffsets
    if (-not $Version) {
        Write-Log "Nepodarilo se extrahovat verzi z vystupu" "ERROR"
        Write-Log "Vystup: $NewOffsets" "ERROR"
        return 1
    }

    Write-Log "Nalezena verze: $Version"

    # Nacist aktualni INI a zkontrolovat zda verze existuje
    $CurrentIni = Get-Content -Path $RdpWrapIniPath -Raw -Encoding UTF8

    if ((Test-VersionExists -Version $Version -IniContent $CurrentIni) -and -not $Force) {
        Write-Log "Verze $Version uz existuje v rdpwrap.ini - zadna akce potreba" "SUCCESS"
        return 0
    }

    if ($Force -and (Test-VersionExists -Version $Version -IniContent $CurrentIni)) {
        Write-Log "Verze $Version uz existuje, ale Force je aktivni - pokracuji" "WARN"
    }

    # Vytvorit zalohu
    $BackupPath = Backup-IniFile

    # Pridat nove offsety
    try {
        Add-OffsetsToIni -NewOffsets $NewOffsets -BackupPath $BackupPath
    }
    catch {
        Write-Log "Chyba pri zapisu do rdpwrap.ini: $_" "ERROR"
        Write-Log "Obnovuji ze zalohy..." "WARN"
        Copy-Item -Path $BackupPath -Destination $RdpWrapIniPath -Force
        return 1
    }

    Write-Log "Aktualizace dokoncena uspesne" "SUCCESS"

    # Restart pokud je pozadovan
    if ($AutoRestart) {
        Write-Log "AutoRestart je aktivni - restartuji pocitac za 30 sekund..." "WARN"
        Write-Log "Pro zruseni spustte: shutdown /a"
        shutdown /r /t 30 /c "RDP Wrapper aktualizovan - restart systemu"
    }
    else {
        Write-Log "Pro aktivaci novych offsetu je potreba restartovat pocitac" "WARN"
    }

    return 0
}

# Spustit
$ExitCode = Main
exit $ExitCode

<#
.SYNOPSIS
    Generates manifest.json for client-side hash-based self-update.

.DESCRIPTION
    Walks the distributed file set rooted at this script's directory, computes
    SHA256 for each file and writes manifest.json next to this script.

    The manifest format consumed by Update-RDPWrap.ps1:

        {
          "version": "2.6.0",
          "generated": "2026-05-28T09:30:00Z",
          "files": [
            { "path": "Update-RDPWrap.ps1",                       "sha256": "...", "size": 27000 },
            { "path": "OffsetFinder/RDPWrapOffsetFinder.exe",     "sha256": "...", "size": 12288 },
            ...
          ]
        }

    Paths in the manifest use forward slashes and are relative to the AutoUpdate
    root. Upload the AutoUpdate folder (including manifest.json) to:
    https://pigman91.github.io/RDP_Patch_Updater/AutoUpdate/

.PARAMETER Root
    AutoUpdate folder to scan (defaults to the directory containing this script).

.PARAMETER Version
    Manifest version string. Defaults to the $ScriptVersion declared inside
    Update-RDPWrap.ps1, or "0.0.0" if it can't be parsed.

.EXAMPLE
    .\Build-Manifest.ps1
#>

param(
    [string]$Root,
    [string]$Version
)

if ([string]::IsNullOrEmpty($Root)) {
    $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$Root = (Resolve-Path $Root).Path

# Files to publish. Anything not in this list is ignored, so logs, backups,
# temp files etc. never leak into the manifest.
$Include = @(
    "Update-RDPWrap.ps1",
    "Install-ScheduledTask.ps1",
    "OffsetFinder/RDPWrapOffsetFinder.exe",
    "OffsetFinder/RDPWrapOffsetFinder_nosymbol.exe",
    "OffsetFinder/dbghelp.dll",
    "OffsetFinder/symsrv.dll",
    "OffsetFinder/Zydis.dll",
    "OffsetFinder/symsrv.yes"
)

# Derive version from Update-RDPWrap.ps1 if not supplied
if ([string]::IsNullOrEmpty($Version)) {
    $ScriptFile = Join-Path $Root "Update-RDPWrap.ps1"
    if (Test-Path $ScriptFile) {
        $Match = Select-String -Path $ScriptFile -Pattern '\$ScriptVersion\s*=\s*"([^"]+)"' | Select-Object -First 1
        if ($Match) {
            $Version = $Match.Matches[0].Groups[1].Value
        }
    }
    if ([string]::IsNullOrEmpty($Version)) { $Version = "0.0.0" }
}

Write-Host "Root:    $Root"
Write-Host "Version: $Version"
Write-Host ""

$Entries = New-Object System.Collections.ArrayList
$Missing = New-Object System.Collections.ArrayList

foreach ($Rel in $Include) {
    $FullPath = Join-Path $Root ($Rel -replace '/', '\')
    if (-not (Test-Path $FullPath)) {
        [void]$Missing.Add($Rel)
        continue
    }
    $Hash = (Get-FileHash -Path $FullPath -Algorithm SHA256).Hash.ToUpper()
    $Size = (Get-Item $FullPath).Length
    [void]$Entries.Add([ordered]@{
        path   = $Rel
        sha256 = $Hash
        size   = $Size
    })
    Write-Host ("  {0,-50} {1}  ({2:N0} B)" -f $Rel, $Hash.Substring(0, 12), $Size)
}

if ($Missing.Count -gt 0) {
    Write-Host ""
    Write-Host "Missing files (not included in manifest):" -ForegroundColor Yellow
    foreach ($M in $Missing) { Write-Host "  - $M" -ForegroundColor Yellow }
}

$Manifest = [ordered]@{
    version   = $Version
    generated = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    files     = $Entries
}

$OutPath = Join-Path $Root "manifest.json"
$Manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $OutPath -Encoding UTF8

Write-Host ""
Write-Host "Manifest written: $OutPath" -ForegroundColor Green
Write-Host "Files in manifest: $($Entries.Count)"

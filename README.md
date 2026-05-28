# RDP Wrapper AutoUpdate

Automatic offset updater for RDP Wrapper that keeps your `rdpwrap.ini` always up-to-date after Windows updates.

## Problem

After every Windows update, the `termsrv.dll` file changes and RDP Wrapper stops working because the offsets in `rdpwrap.ini` no longer match. You have to manually find new offsets or wait for someone to update the INI file.

## Solution

This tool automatically:
1. Detects your current `termsrv.dll` version
2. Generates correct offsets using RDPWrapOffsetFinder
3. Adds them to `rdpwrap.ini` if not already present
4. Optionally restarts your PC to apply changes

The scheduled task runs at every system startup, so your RDP Wrapper stays working even after Windows updates.

## Download

Download the latest release from the [Releases](https://github.com/pigman91/RDP_Patch_Updater/releases) page.

## Installation Options

### Option 1: Install-RDPWrapperComplete.exe (Recommended for new users)

**Use this if you DON'T have RDP Wrapper installed yet.**

This complete installer will:
- Download and install RDP Wrapper (RDPWInst.exe, RDPConf.exe, RDPCheck.exe)
- Download and install AutoUpdate components (OffsetFinder, scripts)
- Create a scheduled task that runs at every system startup
- Run the first offset check immediately
- Copy RDPConf.exe to the installation folder for easy access
- Automatically restart PC if new offsets are added (with 30 second warning)

### Option 2: Install-RDPAutoUpdate.exe (For existing RDP Wrapper users)

**Use this if you ALREADY have RDP Wrapper installed and working.**

This AutoUpdate-only installer will:
- Download and install AutoUpdate components (OffsetFinder, scripts)
- Create a scheduled task that runs at every system startup
- Run the first offset check immediately
- Automatically restart PC if new offsets are added (with 30 second warning)

## Requirements

- Windows 10/11
- Administrator privileges
- .NET Framework 4.x
- Internet connection (for downloading files during installation)

## How It Works

1. **Scheduled Task** "RDP Wrapper Auto Update" runs at system startup
2. **Update-RDPWrap.ps1** first downloads `manifest.json` and refreshes any of its own files (script, OffsetFinder, DLLs) whose SHA256 no longer matches — so binaries stay current, not just the script (see [Self-update](#self-update) below)
3. The script executes RDPWrapOffsetFinder.exe against your termsrv.dll
4. The tool extracts the version number (e.g., 10.0.26100.2605)
5. If this version is not found in `rdpwrap.ini`, the new offsets are added
6. A backup of the original INI is created before any changes
7. If AutoRestart is enabled, PC restarts in 30 seconds to apply changes
8. You can cancel the restart by running `shutdown /a`

## Self-update

Starting with v2.6.0 the updater pulls a `manifest.json` from
`https://pigman91.github.io/RDP_Patch_Updater/AutoUpdate/manifest.json` listing
every distributed file with its SHA256 hash. On each run the client compares the
listed hashes against its local files and downloads only the ones that differ.
Each downloaded file is hash-verified before it replaces the local copy.

This is the difference from v2.5.x, where the self-updater only replaced the
`.ps1` and never touched the OffsetFinder binaries — meaning a Windows update
that broke the byte patterns in OffsetFinder could never be fixed remotely.

### Publishing a new release

Whenever any file inside `AutoUpdate/` changes (new OffsetFinder, script fix,
…), regenerate the manifest before uploading:

```powershell
& .\AutoUpdate\Build-Manifest.ps1
```

Then upload the full `AutoUpdate/` folder — **including the regenerated
`manifest.json`** — to the GitHub Pages mirror. If you upload new binaries but
forget to refresh the manifest, clients will keep the old files.

### Useful script parameters

```powershell
# Run only the manifest update step, skip OffsetFinder + INI patch
.\Update-RDPWrap.ps1 -ManifestOnly

# Point at an internal mirror or a test server instead of GitHub Pages
.\Update-RDPWrap.ps1 -UpdateBaseUrl "http://mirror.local/rdp"
```

## Installation Directory Structure

```
C:\Program Files\RDP Wrapper\
├── rdpwrap.ini              # RDP Wrapper configuration (offsets)
├── rdpwrap.dll              # RDP Wrapper library
├── RDPConf.exe              # Configuration GUI tool
└── AutoUpdate\
    ├── Install-ScheduledTask.ps1   # Scheduled task installer
    ├── Update-RDPWrap.ps1          # Main update script
    ├── Update-RDPWrap.log          # Log file
    ├── .last_termsrv_hash          # Cached hash so unchanged DLLs skip the check
    ├── .last_script_version        # Version marker, forces full check after upgrade
    ├── Backups\                    # INI backups before changes
    │   └── rdpwrap_YYYYMMDD_HHMMSS.ini
    └── OffsetFinder\
        ├── RDPWrapOffsetFinder.exe         # Offset finder (with symbols)
        ├── RDPWrapOffsetFinder_nosymbol.exe # Offset finder (without symbols)
        ├── dbghelp.dll
        ├── symsrv.dll
        ├── Zydis.dll
        ├── symsrv.yes
        └── sym\                            # PDB cache (auto-managed)
```

### Repository / GitHub Pages layout

```
AutoUpdate/
├── Install-ScheduledTask.ps1
├── Update-RDPWrap.ps1
├── Build-Manifest.ps1       # NOT distributed to clients — used to publish releases
├── manifest.json            # SHA256 of every distributed file; clients fetch this first
└── OffsetFinder/
    ├── RDPWrapOffsetFinder.exe
    ├── RDPWrapOffsetFinder_nosymbol.exe
    ├── dbghelp.dll
    ├── symsrv.dll
    ├── Zydis.dll
    └── symsrv.yes
```

## Manual Usage

You can also run the update manually:

```powershell
# Check and update offsets (no restart)
& "C:\Program Files\RDP Wrapper\AutoUpdate\Update-RDPWrap.ps1"

# Check and update offsets (auto restart if needed)
& "C:\Program Files\RDP Wrapper\AutoUpdate\Update-RDPWrap.ps1" -AutoRestart

# Force update even if version exists
& "C:\Program Files\RDP Wrapper\AutoUpdate\Update-RDPWrap.ps1" -Force
```

## Uninstallation

### Remove scheduled task only:
```powershell
& "C:\Program Files\RDP Wrapper\AutoUpdate\Install-ScheduledTask.ps1" -Uninstall
```

### Complete removal:
1. Run the command above to remove the scheduled task
2. Delete the folder `C:\Program Files\RDP Wrapper\AutoUpdate`
3. (Optional) Uninstall RDP Wrapper using its uninstall.bat

## Troubleshooting

### Check the log file
```
C:\Program Files\RDP Wrapper\AutoUpdate\Update-RDPWrap.log
```

### Common issues:
- **"RDPWrapOffsetFinder.exe not found"** - Reinstall using the installer
- **"rdpwrap.ini not found"** - RDP Wrapper is not installed, use Complete installer
- **"OffsetFinder could not find all offsets" / "DefPolicyPatch patten not found"** - Microsoft changed an internal byte pattern in this `termsrv.dll`. Publish a refreshed `RDPWrapOffsetFinder.exe` in the GitHub Pages mirror (with a regenerated `manifest.json`); clients will pick it up on the next scheduled run.
- **"Too many users connected" on the RDP login screen** - Same root cause: your current `termsrv.dll` version isn't in `rdpwrap.ini`, so the wrapper isn't active and Windows enforces its default single-session limit. The next successful updater run that adds the new version + a reboot fixes it.

## Credits

- [RDP Wrapper Library](https://github.com/stascorp/rdpwrap) by stascorp
- [RDPWrapOffsetFinder](https://github.com/llccd/RDPWrapOffsetFinder) by llccd

## License

MIT License

## Disclaimer

This tool is provided as-is for educational purposes. Use at your own risk. The author is not responsible for any damage caused by using this software. Make sure you comply with Microsoft's licensing terms.

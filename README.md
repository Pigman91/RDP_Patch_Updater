# RDP Wrapper AutoUpdate

Automatic offset updater for RDP Wrapper that keeps your `rdpwrap.ini` always up-to-date after Windows updates.

[![PayPal Donate](https://img.shields.io/badge/PayPal-Donate-blue.svg?logo=paypal)](https://www.paypal.com/donate/?hosted_button_id=J8DWRTBTMZKVE)

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
2. **Update-RDPWrap.ps1** executes RDPWrapOffsetFinder.exe against your termsrv.dll
3. The tool extracts the version number (e.g., 10.0.26100.2605)
4. If this version is not found in `rdpwrap.ini`, the new offsets are added
5. A backup of the original INI is created before any changes
6. If AutoRestart is enabled, PC restarts in 30 seconds to apply changes
7. You can cancel the restart by running `shutdown /a`

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
    ├── Backups\                    # INI backups before changes
    │   └── rdpwrap_YYYYMMDD_HHMMSS.ini
    └── OffsetFinder\
        ├── RDPWrapOffsetFinder.exe         # Offset finder (with symbols)
        ├── RDPWrapOffsetFinder_nosymbol.exe # Offset finder (without symbols)
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
- **"OffsetFinder could not find all offsets"** - Your Windows version might not be supported yet

## Support

If you find this tool useful, consider buying me a coffee:

[![PayPal Donate](https://img.shields.io/badge/PayPal-Donate-blue.svg?logo=paypal)](https://www.paypal.com/donate/?hosted_button_id=J8DWRTBTMZKVE)

## Credits

- [RDP Wrapper Library](https://github.com/stascorp/rdpwrap) by stascorp
- [RDPWrapOffsetFinder](https://github.com/llccd/RDPWrapOffsetFinder) by llccd

## License

MIT License

## Disclaimer

This tool is provided as-is for educational purposes. Use at your own risk. The author is not responsible for any damage caused by using this software. Make sure you comply with Microsoft's licensing terms.

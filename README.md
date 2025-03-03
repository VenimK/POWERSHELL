# PowerShell Scripts Collection

A collection of useful PowerShell scripts for various system administration, automation, and utility tasks.

## 🚀 Scripts Overview

### System Information & Management
- **CPUINFO.ps1**: Displays detailed CPU information and specifications
- **Drivers.ps1**: Manages and displays system drivers information
- **battery.ps1**: Shows battery status and health information
- **bios.ps1**: Retrieves BIOS information and settings
- **TPMByPass.ps1**: Helps bypass TPM requirements for Windows installations

### System Maintenance
- **clean.ps1**: System cleanup utility
- **cleandisk.ps1**: Disk cleanup and optimization tool
- **dism.ps1**: Windows system image maintenance and repair
- **WUA.ps1**: Windows Update Assistant automation
- **updater.ps1**: System update automation tool

### WiFi & Network Tools
- **QRMostUsedWifi.ps1**: Generates QR codes for your most frequently used WiFi networks, including network name and password
- **QWIFI.ps1**: Quick WiFi connection utility
- **WifiHistory.ps1**: Shows WiFi connection history with connection counts
- **ExportWLAN.ps1**: Exports WiFi network profiles for backup or transfer
- **ImportWLAN.ps1**: Imports WiFi network profiles from backup

### Browser Tools
- **fetchchrome.ps1**: Chrome browser data extraction utility
- **fetchchromebookmarks.ps1**: Exports Chrome bookmarks to a file
- **fetchedge.ps1**: Microsoft Edge data management and backup

### Software Management
- **ExportApps.ps1**: Exports list of installed applications to a file
- **ExportWinget.ps1**: Exports winget package list for backup
- **ImportWinget.ps1**: Imports and installs winget packages from list
- **install-winget.ps1**: Automated winget package manager installation
- **installed.ps1**: Shows detailed installed software information

### Document Processing
- **EANGenerator.ps1**: Generates and inserts EAN-13 barcodes into Word documents, supporting multiple barcodes per document

### System Tweaks & Utilities
- **tweak.ps1**: System optimization and performance tweaks
- **signwin.ps1**: Windows signing utility for scripts and executables
- **list.ps1**: Lists system information and configurations

### Notifications
- **noti.ps1**: Custom notification system
- **notid.ps1**: Notification daemon/service for background alerts

## 🔧 Usage

1. Make sure you have PowerShell installed on your system
2. Clone this repository or download the scripts you need
3. Run PowerShell as Administrator for scripts that require elevated privileges
4. Execute scripts using one of these methods:
   ```powershell
   # Method 1: Direct execution with bypass
   powershell -ExecutionPolicy Bypass -File script_name.ps1

   # Method 2: After setting RemoteSigned policy
   Set-ExecutionPolicy RemoteSigned
   .\script_name.ps1
   ```

## ⚠️ Requirements

- Windows 10/11
- PowerShell 5.1 or later
- Administrator privileges for some scripts
- Internet connection for scripts that download resources

## 🔒 Security Note

Some scripts require administrator privileges. Always review scripts before running them and ensure they come from a trusted source.

## 👤 Author

VenimK (techmusiclover@outlook.be)

## 📝 License

These scripts are provided as-is under the MIT License. Feel free to modify and distribute them while maintaining attribution.

# Architecture Overview

The application is split into two components:

## 1. Hardware Monitor Service (`hardware_monitor_service.py`)
- **Purpose**: Windows Service that runs with SYSTEM privileges
- **Privileges**: Runs automatically at system startup without UAC prompts
- **Responsibilities**:
  - Hardware monitoring using LibreHardwareMonitorLib
  - Data collection and processing
  - Named pipe server for IPC communication
- **Communication**: Serves data via named pipe `\\.\pipe\TrayHardwareMonitor`

## 2. Tray Serial Monitor Client (`tray_serial_monitor_client.py`)
- **Purpose**: User-space GUI application
- **Privileges**: Runs in normal user context (no admin required)
- **Responsibilities**:
  - System tray interface
  - Serial communication
  **Communication**: Connects to service via named pipe

# Installation

**No Python installation required for end users!**

## For End Users (Pre-built Installer)
1. Download `TraySerialMonitor_Standalone_Setup.exe` from the releases or from `installer_output/` (if building from source)
2. Right-click and "Run as administrator"
3. Follow the installation wizard
4. Choose installation options:
   - **Start client automatically when Windows starts** (recommended)
   - ☐ Create desktop icon (optional)

## For Developers (Building from Source)
If you're building from source, you'll need to generate the installer first:
1. Run `build_complete_installer.bat` to create the installer
2. The installer will be generated in `installer_output/TraySerialMonitor_Standalone_Setup.exe`
3. Follow the end user installation steps above

The installer will:
- Install standalone executables (no Python required)
- Install and start the Windows service automatically
- Create Start Menu shortcuts
- Add entry to Windows Programs & Features for easy uninstallation

## Prerequisites
- Windows 10/11 (64-bit)
- Administrator privileges for service installation
- **For end users**: No additional requirements
- **For developers**: Python 3.x with required packages

## Usage

## Starting the Application
After installation:
1. The service starts automatically at system boot
2. The client application starts automatically when you log in
3. Look for the system tray icon to access the interface

## Manual Control
- **Start Service**: `net start TrayHardwareMonitor`
- **Stop Service**: `net stop TrayHardwareMonitor`
- **Service Status**: Check Windows Services (`services.msc`)

# Uninstallation

Use the windows remove program feature to uninstall the software.

You can also run the uninstall script as Administrator:
```cmd
python uninstall_service.py
```

This will:
- Stop and remove the Windows service
- Remove startup shortcuts
- Clean up service files

## File Structure

### Source Files (Git Repository)
```
front-panel-tray/
├── hardware_monitor_service.py    # Windows service component
├── tray_serial_monitor_client.py  # User interface client
├── install_service.py             # Service installation script
├── uninstall_service.py           # Service removal script
├── build_complete_installer.bat   # Complete build automation script
├── build_executables.py           # Python script for building executables
├── service_installer.iss          # Inno Setup installer script
├── icon.ico                       # Application icon
├── LibreHardwareMonitorLib.dll    # Hardware monitoring library
├── LibreHardwareMonitorLib.sys    # Hardware monitoring driver
└── README.md                      # This documentation
```

### Generated Build Artifacts (Not in Git)
**Note**: The following directories and files are created by the build process and are not included in the git repository:

```
front-panel-tray/
├── exe/                           # Built executables directory (generated)
│   ├── TrayHardwareMonitorService.exe  # Service executable
│   ├── TraySerialMonitorClient.exe     # Client executable
│   ├── InstallService.exe              # Service installer executable
│   ├── UninstallService.exe            # Service uninstaller executable
│   ├── icon.ico                        # Application icon (copied)
│   ├── LibreHardwareMonitorLib.dll     # Hardware monitoring library (copied)
│   ├── LibreHardwareMonitorLib.sys     # Hardware monitoring driver (copied)
│   └── README.md                       # Documentation (copied)
└── installer_output/              # Generated installer directory
    └── TraySerialMonitor_Standalone_Setup.exe  # Complete installer package
```

### Build Process File Generation

The build process creates files in the following sequence:

1. **`build_executables.py`** creates the `exe/` directory and generates:
   - Standalone executables using PyInstaller
   - Copies required dependencies (DLLs, icons, documentation)
   - Creates service installer/uninstaller executables

2. **`service_installer.iss`** (via Inno Setup) creates the `installer_output/` directory and generates:
   - Complete Windows installer package (`TraySerialMonitor_Standalone_Setup.exe`)
   - Packages all executables and dependencies into a single installer

3. **`build_complete_installer.bat`** orchestrates the entire process:
   - Runs the Python build script
   - Executes Inno Setup compilation
   - Creates the final distributable installer

**To build from source:**
```cmd
# Run the complete build process
build_complete_installer.bat

# Or build executables only
python build_executables.py
```
## Troubleshooting

### Service Won't Start
1. Check Windows Event Viewer for service errors
2. Verify all dependencies are present
3. Ensure LibreHardwareMonitorLib files are accessible
4. Run `python hardware_monitor_service.py debug` for detailed output

### Client Can't Connect to Service
1. Verify service is running: `net start TrayHardwareMonitor`
2. Check named pipe permissions
3. Restart both service and client
4. Review client application logs

### Hardware Data Not Available
1. Ensure service is running with SYSTEM privileges
2. Check if LibreHardwareMonitorLib.sys driver is loaded
3. Verify hardware monitoring permissions
4. Test with original monolithic version for comparison

## Technical Details

### Inter-Process Communication
- **Method**: Named Pipes (`\\.\pipe\TrayHardwareMonitor`)
- **Protocol**: JSON-formatted messages
- **Security**: Local access only

### Service Management
- **Service Name**: `TrayHardwareMonitor`
- **Display Name**: `Tray Hardware Monitor Service`
- **Start Type**: Automatic
- **Account**: Local System

### Dependencies
- `pywin32`: Windows service framework and named pipes
- `pyserial`: Serial communication
- `pystray`: System tray interface
- `LibreHardwareMonitorLib`: Hardware monitoring
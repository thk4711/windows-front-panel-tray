"""
Build standalone executables for the service-based Tray Serial Monitor
This script uses PyInstaller to create .exe files that don't require Python installation
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def run_command(cmd, description):
    """Run a command and handle errors"""
    print(f"\n{description}...")
    print(f"Running: {cmd}")

    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)

    if result.returncode == 0:
        print(f"{description} completed successfully")
        if result.stdout:
            print("Output:", result.stdout.strip())
    else:
        print(f"{description} failed!")
        print("Error:", result.stderr.strip())
        return False

    return True

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        result = subprocess.run(['python', '-m', 'PyInstaller', '--version'], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"PyInstaller found: {result.stdout.strip()}")
            return True
    except FileNotFoundError:
        pass

    print("PyInstaller not found. Installing...")
    return run_command('python -m pip install pyinstaller', 'Installing PyInstaller')

def create_service_spec():
    """Create PyInstaller spec file for the service"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['hardware_monitor_service.py'],
    pathex=[],
    binaries=[
        ('LibreHardwareMonitorLib.dll', '.'),
        ('LibreHardwareMonitorLib.sys', '.'),
    ],
    datas=[
        ('icon.ico', '.'),
    ],
    hiddenimports=[
        'win32serviceutil',
        'win32service',
        'win32event',
        'win32pipe',
        'win32file',
        'win32api',
        'win32con',
        'pywintypes',
        'servicemanager',
        'win32security',
        'win32process',
        'winerror',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywin32_system32',
        'win32timezone',
        'win32clipboard',
        'win32gui',
        'win32pdh',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='TrayHardwareMonitorService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir='C:\\TrayTemp',
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
'''

    with open('service.spec', 'w') as f:
        f.write(spec_content)
    print("Created service.spec")

def create_client_spec():
    """Create PyInstaller spec file for the client"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['tray_serial_monitor_client.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon.ico', '.'),
    ],
    hiddenimports=[
        'pystray',
        'PIL',
        'PIL.Image',
        'serial',
        'win32file',
        'win32pipe',
        'win32con',
        'win32api',
        'win32event',
        'pywintypes',
        'json',
        'threading',
        'time',
        'tkinter',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        # Additional win32 modules for named pipe communication
        'win32security',
        'win32process',
        'winerror',
        'win32service',
        'win32serviceutil',
        'servicemanager',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywin32_system32',
        # Additional modules that might be needed
        'win32clipboard',
        'win32gui',
        'win32pdh',
        'win32timezone',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='TraySerialMonitorClient',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir='C:\\TrayTemp',
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
'''

    with open('client.spec', 'w') as f:
        f.write(spec_content)
    print("Created client.spec")

def create_installer_spec():
    """Create PyInstaller spec file for the installer"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['install_service.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon.ico', '.'),
    ],
    hiddenimports=[
        'win32serviceutil',
        'win32service',
        'win32api',
        'win32con',
        'win32com.client',
        'subprocess',
        'os',
        'sys',
        'ctypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='InstallService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir='C:\\TrayTemp',
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
'''

    with open('installer.spec', 'w') as f:
        f.write(spec_content)
    print("Created installer.spec")

def create_uninstaller_spec():
    """Create PyInstaller spec file for the uninstaller"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['uninstall_service.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32serviceutil',
        'win32service',
        'win32api',
        'subprocess',
        'os',
        'sys',
        'ctypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='UninstallService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir='C:\\TrayTemp',
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
'''

    with open('uninstaller.spec', 'w') as f:
        f.write(spec_content)
    print("Created uninstaller.spec")

def build_executables():
    """Build all executables"""
    executables = [
        ('service.spec', 'Building Windows Service executable'),
        ('client.spec', 'Building Client GUI executable'),
        ('installer.spec', 'Building Service Installer executable'),
        ('uninstaller.spec', 'Building Service Uninstaller executable'),
    ]

    # Create dist directory if it doesn't exist
    os.makedirs('dist', exist_ok=True)

    for spec_file, description in executables:
        if not run_command(f'python -m PyInstaller --clean {spec_file}', description):
            return False

    return True

def organize_executables():
    """Organize built executables into a clean structure"""
    print("\nOrganizing executables...")

    # Create exe directory
    exe_dir = Path('exe')
    if exe_dir.exists():
        shutil.rmtree(exe_dir)
    exe_dir.mkdir()

    # Copy executables
    executables = [
        ('TrayHardwareMonitorService.exe', 'TrayHardwareMonitorService.exe'),
        ('TraySerialMonitorClient.exe', 'TraySerialMonitorClient.exe'),
        ('InstallService.exe', 'InstallService.exe'),
        ('UninstallService.exe', 'UninstallService.exe'),
    ]

    for exe_name, target_name in executables:
        src = Path('dist') / exe_name
        dst = exe_dir / target_name

        if src.exists():
            shutil.copy2(src, dst)
            print(f"[OK] Copied {exe_name} -> exe/{target_name}")
        else:
            print(f"[ERROR] Missing {exe_name}")
            return False

    # Copy additional files
    additional_files = [
        'icon.ico',
        'LibreHardwareMonitorLib.dll',
        'LibreHardwareMonitorLib.sys',
        'README.md',
        'QUICK_START.md',
        'INSTALLER_GUIDE.md',
    ]

    for file_name in additional_files:
        src = Path(file_name)
        dst = exe_dir / file_name

        if src.exists():
            shutil.copy2(src, dst)
            print(f"[OK] Copied {file_name}")
        else:
            print(f"[WARNING] Missing {file_name} (optional)")

    return True

def cleanup_build_files():
    """Clean up build artifacts"""
    print("\nCleaning up build files...")

    cleanup_items = [
        'build',
        'dist',
        '*.spec',
        '__pycache__',
    ]

    for item in cleanup_items:
        if '*' in item:
            # Handle wildcards
            import glob
            for path in glob.glob(item):
                if os.path.isfile(path):
                    os.remove(path)
                    print(f"[OK] Removed {path}")
        else:
            path = Path(item)
            if path.exists():
                if path.is_dir():
                    shutil.rmtree(path)
                else:
                    path.unlink()
                print(f"[OK] Removed {path}")

def main():
    """Main build process"""
    print("Building Standalone Executables for Tray Serial Monitor Service")
    print("=" * 70)

    # Create temp directory if it doesn't exist (needed for PyInstaller build)
    temp_dir = Path('C:\\TrayTemp')
    if not temp_dir.exists():
        try:
            temp_dir.mkdir(parents=True, exist_ok=True)
            print("[OK] Created C:\\TrayTemp directory for PyInstaller")
        except PermissionError:
            print("[WARNING] Could not create C:\\TrayTemp (permission denied)")
            print("   PyInstaller will attempt to create it during runtime")

    # Check PyInstaller
    if not check_pyinstaller():
        return False

    # Create spec files
    print("\nCreating PyInstaller spec files...")
    create_service_spec()
    create_client_spec()
    create_installer_spec()
    create_uninstaller_spec()

    # Build executables
    print("\nBuilding executables...")
    if not build_executables():
        print("\n[ERROR] Build failed!")
        return False

    # Organize files
    if not organize_executables():
        print("\n[ERROR] Failed to organize executables!")
        return False

    # Cleanup
    cleanup_build_files()

    print("\n" + "=" * 70)
    print("SUCCESS! Standalone executables created in 'exe' directory:")
    print("   exe/")
    print("   ├── TrayHardwareMonitorService.exe  (Windows Service)")
    print("   ├── TraySerialMonitorClient.exe     (GUI Client with ESP32 Auto-Detection)")
    print("   ├── InstallService.exe              (Service Installer)")
    print("   ├── UninstallService.exe            (Service Uninstaller)")
    print("   ├── icon.ico                        (Application Icon)")
    print("   ├── LibreHardwareMonitorLib.dll     (Hardware Library)")
    print("   ├── LibreHardwareMonitorLib.sys     (Hardware Driver)")
    print("   └── Documentation files...")
    print("\nNo Python installation required for end users!")
    print("Ready to create Windows installer with standalone executables")

    return True

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)

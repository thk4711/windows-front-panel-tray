"""
Silent Service Uninstallation Script
This script removes the hardware monitoring service without user interaction.
Designed for use with automated installers.
"""

import os
import sys
import subprocess

def is_admin():
    """Check if running as administrator"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def uninstall_service():
    """Uninstall the hardware monitoring service"""
    print("Uninstalling Hardware Monitor Service...")
    
    # When running as a PyInstaller executable, __file__ points to the temp directory
    # We need to find the actual installation directory where the service exe is located
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Check if we're running from a PyInstaller temp directory
    if "_MEI" in current_dir:
        # We're running from PyInstaller temp directory, need to find the real installation directory
        # Try to get the directory where the uninstaller executable is located
        import sys
        if hasattr(sys, '_MEIPASS'):
            # Get the directory containing the uninstaller executable
            uninstaller_dir = os.path.dirname(sys.executable)
            print(f"Detected PyInstaller environment. Looking in uninstaller directory: {uninstaller_dir}")
            current_dir = uninstaller_dir
    
    # For standalone executable, look for the service exe file
    service_exe = os.path.join(current_dir, "TrayHardwareMonitorService.exe")
    
    if not os.path.exists(service_exe):
        print(f"WARNING: Service executable not found at {service_exe}")
        print(f"Current directory: {current_dir}")
        print("Available files:")
        try:
            for file in os.listdir(current_dir):
                print(f"  - {file}")
        except:
            pass
        return True  # Consider it already uninstalled
    
    try:
        # Stop the service first
        print("Stopping service...")
        cmd = [service_exe, "stop"]
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=current_dir, timeout=30)
        
        if result.returncode == 0:
            print("✓ Service stopped successfully")
        else:
            print(f"⚠ Service stop result: {result.stderr}")
        
        # Uninstall the service
        print("Removing service...")
        cmd = [service_exe, "remove"]
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=current_dir, timeout=30)
        
        if result.returncode == 0:
            print("✓ Service removed successfully")
            return True
        else:
            print(f"✗ Failed to remove service: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("✗ Service uninstallation timed out")
        return False
    except Exception as e:
        print(f"✗ Error uninstalling service: {e}")
        return False


def main():
    """Main uninstallation function - silent mode"""
    if not is_admin():
        print("ERROR: This script must be run as Administrator!")
        sys.exit(1)
    
    print("Uninstalling service-based Tray Serial Monitor (silent mode)...")
    
    # Uninstall service
    if not uninstall_service():
        print("✗ Service uninstallation failed!")
        sys.exit(1)
    
    print("✓ Uninstallation completed successfully!")
    return True

if __name__ == "__main__":
    main()

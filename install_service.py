"""
Silent Service Installation Script
This script installs the hardware monitoring service without user interaction.
Designed for use with automated installers.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def is_admin():
    """Check if running as administrator"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def configure_service_recovery(hscm, hs):
    """Configure service recovery options for automatic restart on failure"""
    print("Configuring service recovery options...")
    
    try:
        import win32service
        import ctypes
        from ctypes import wintypes
        
        # Define recovery actions structure
        class SERVICE_FAILURE_ACTIONS(ctypes.Structure):
            _fields_ = [
                ('dwResetPeriod', wintypes.DWORD),
                ('lpRebootMsg', wintypes.LPWSTR),
                ('lpCommand', wintypes.LPWSTR),
                ('cActions', wintypes.DWORD),
                ('lpsaActions', ctypes.POINTER(ctypes.c_void_p))
            ]
        
        class SC_ACTION(ctypes.Structure):
            _fields_ = [
                ('Type', wintypes.DWORD),
                ('Delay', wintypes.DWORD)
            ]
        
        # Recovery action types
        SC_ACTION_NONE = 0
        SC_ACTION_RESTART = 1
        SC_ACTION_REBOOT = 2
        SC_ACTION_RUN_COMMAND = 3
        
        # Create recovery actions array
        actions = (SC_ACTION * 3)()
        actions[0] = SC_ACTION(SC_ACTION_RESTART, 60000)    # Restart after 1 minute
        actions[1] = SC_ACTION(SC_ACTION_RESTART, 120000)   # Restart after 2 minutes
        actions[2] = SC_ACTION(SC_ACTION_RESTART, 300000)   # Restart after 5 minutes
        
        # Create failure actions structure
        failure_actions = SERVICE_FAILURE_ACTIONS()
        failure_actions.dwResetPeriod = 86400  # Reset failure count after 24 hours
        failure_actions.lpRebootMsg = None
        failure_actions.lpCommand = None
        failure_actions.cActions = 3
        failure_actions.lpsaActions = ctypes.cast(actions, ctypes.POINTER(ctypes.c_void_p))
        
        # Use ChangeServiceConfig2 to set recovery options
        SERVICE_CONFIG_FAILURE_ACTIONS = 2
        advapi32 = ctypes.windll.advapi32
        
        result = advapi32.ChangeServiceConfig2W(
            hs,
            SERVICE_CONFIG_FAILURE_ACTIONS,
            ctypes.byref(failure_actions)
        )
        
        if result:
            print("✓ Service recovery options configured successfully")
            print("  - First failure: Restart after 1 minute")
            print("  - Second failure: Restart after 2 minutes")
            print("  - Subsequent failures: Restart after 5 minutes")
            print("  - Reset failure count after 24 hours")
        else:
            error_code = ctypes.windll.kernel32.GetLastError()
            print(f"✗ Failed to configure service recovery options (Error: {error_code})")
            
            # Fallback: Use sc command
            print("Attempting fallback configuration using sc command...")
            try:
                cmd = [
                    'sc', 'failure', 'TrayHardwareMonitor',
                    'reset=', '86400',  # Reset after 24 hours
                    'actions=', 'restart/60000/restart/120000/restart/300000'
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
                if result.returncode == 0:
                    print("✓ Service recovery configured using sc command")
                else:
                    print(f"✗ sc command failed: {result.stderr}")
            except Exception as e:
                print(f"✗ Fallback configuration failed: {e}")
        
    except Exception as e:
        print(f"✗ Error configuring service recovery: {e}")
        # Try simple sc command as last resort
        try:
            print("Attempting simple sc command configuration...")
            cmd = ['sc', 'failure', 'TrayHardwareMonitor', 'reset=', '86400', 'actions=', 'restart/60000']
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                print("✓ Basic service recovery configured")
            else:
                print(f"✗ Simple sc command failed: {result.stderr}")
        except Exception as e2:
            print(f"✗ All recovery configuration methods failed: {e2}")

def install_service():
    """Install the hardware monitoring service"""
    print("Installing Hardware Monitor Service...")
    
    # When running as a PyInstaller executable, __file__ points to the temp directory
    # We need to find the actual installation directory where the service exe is located
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Check if we're running from a PyInstaller temp directory
    if "_MEI" in current_dir:
        # We're running from PyInstaller temp directory, need to find the real installation directory
        # Try to get the directory where the installer executable is located
        import sys
        if hasattr(sys, '_MEIPASS'):
            # Get the directory containing the installer executable
            installer_dir = os.path.dirname(sys.executable)
            print(f"Detected PyInstaller environment. Looking in installer directory: {installer_dir}")
            current_dir = installer_dir
    
    # For standalone executable, look for the service exe file
    service_exe = os.path.join(current_dir, "TrayHardwareMonitorService.exe")
    
    if not os.path.exists(service_exe):
        print(f"ERROR: Service executable not found at {service_exe}")
        print(f"Current directory: {current_dir}")
        print("Available files:")
        try:
            for file in os.listdir(current_dir):
                print(f"  - {file}")
        except:
            pass
        return False
    
    try:
        # Install the service using the standalone executable
        cmd = [service_exe, "install"]
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=current_dir, timeout=30)
        
        if result.returncode == 0:
            print("✓ Service installed successfully")
            
            # Set service to automatic startup
            print("Configuring service for automatic startup...")
            try:
                import win32service
                import win32serviceutil
                
                # Open service manager
                hscm = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_ALL_ACCESS)
                if hscm:
                    # Open the service
                    hs = win32service.OpenService(hscm, "TrayHardwareMonitor", win32service.SERVICE_ALL_ACCESS)
                    if hs:
                        # Change service to automatic startup
                        win32service.ChangeServiceConfig(
                            hs,
                            win32service.SERVICE_NO_CHANGE,  # dwServiceType
                            win32service.SERVICE_AUTO_START,  # dwStartType - this is the key change
                            win32service.SERVICE_NO_CHANGE,  # dwErrorControl
                            None,  # lpBinaryPathName
                            None,  # lpLoadOrderGroup
                            0,     # lpdwTagId
                            None,  # lpDependencies
                            None,  # lpServiceStartName
                            None,  # lpPassword
                            None   # lpDisplayName
                        )
                        print("✓ Service configured for automatic startup")
                        
                        # Configure service recovery options
                        configure_service_recovery(hscm, hs)
                        
                        win32service.CloseServiceHandle(hs)
                        
                    else:
                        print("✗ Failed to open service for configuration")
                    win32service.CloseServiceHandle(hscm)
                else:
                    print("✗ Failed to open service manager")
            except Exception as e:
                print(f"✗ Error configuring service startup: {e}")
                # Continue anyway, service is installed
            
            # Start the service
            print("Starting service...")
            cmd = [service_exe, "start"]
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=current_dir, timeout=30)
            
            if result.returncode == 0:
                print("✓ Service started successfully")
                return True
            else:
                print(f"✗ Failed to start service: {result.stderr}")
                return False
        else:
            print(f"✗ Failed to install service: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("✗ Service installation timed out")
        return False
    except Exception as e:
        print(f"✗ Error installing service: {e}")
        return False


def main():
    """Main installation function - silent mode"""
    if not is_admin():
        print("ERROR: This script must be run as Administrator!")
        sys.exit(1)
    
    print("Installing service-based Tray Serial Monitor (silent mode)...")
    
    # Install service
    if not install_service():
        print("✗ Service installation failed!")
        sys.exit(1)
    
    print("✓ Installation completed successfully!")
    return True

if __name__ == "__main__":
    main()

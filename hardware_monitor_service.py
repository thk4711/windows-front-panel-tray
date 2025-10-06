"""
Hardware Monitor Service - Windows Service Component
This service runs with SYSTEM privileges and handles hardware monitoring.
It exposes data via named pipe for the tray client to consume.
"""

import json
import threading
import time
import sys
import os
from datetime import datetime
import ctypes
import win32serviceutil
import win32service
import win32event
import servicemanager
import win32pipe
import win32file
import pywintypes

import psutil

# For master volume
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# For CPU temperature via LibreHardwareMonitor
import clr

# Monkey-patch comtypes destructor to suppress access violations on release
import comtypes
from comtypes._post_coinit.unknwn import _compointer_base
_orig_del = _compointer_base.__del__
def _safe_del(self):
    try:
        _orig_del(self)
    except Exception:
        pass
_compointer_base.__del__ = _safe_del


class WindowsCPUMonitor:
    """Simple psutil-based CPU monitoring"""
    
    def __init__(self):
        self.initialized = False
        self.last_cpu_value = 0
        
    def initialize(self):
        """Initialize psutil CPU monitoring"""
        try:
            servicemanager.LogInfoMsg("Initializing psutil CPU monitoring...")
            
            # Prime psutil CPU counter
            psutil.cpu_percent(interval=None)
            
            self.initialized = True
            servicemanager.LogInfoMsg("psutil CPU monitoring initialized successfully")
            return True
            
        except Exception as e:
            servicemanager.LogErrorMsg(f"Failed to initialize psutil CPU monitoring: {e}")
            return False
    
    def get_cpu_usage(self):
        """Get CPU usage using standard psutil"""
        try:
            # Use standard psutil with 0.3 second interval (original implementation)
            cpu_percent = psutil.cpu_percent(interval=0.3)
            self.last_cpu_value = int(round(cpu_percent))
            return self.last_cpu_value
                
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error in psutil CPU monitoring: {e}")
            return self.last_cpu_value
    
    def cleanup(self):
        """Clean up CPU monitor resources"""
        try:
            self.initialized = False
            servicemanager.LogInfoMsg("psutil CPU monitor cleaned up")
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error cleaning up psutil CPU monitor: {e}")


class HardwareMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "TrayHardwareMonitor"
    _svc_display_name_ = "Tray Hardware Monitor Service"
    _svc_description_ = "Hardware monitoring service for Tray Serial Monitor"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.stop_event = threading.Event()
        self.pipe_name = r'\\.\pipe\TrayHardwareMonitor'
        self.current_data = {}
        self.data_lock = threading.Lock()
        self.computer = None
        self.Hardware = None  # Store Hardware module reference
        self.cpu_monitor = WindowsCPUMonitor()  # Simple psutil CPU monitor
        
        # Don't initialize LibreHardwareMonitor here - do it in SvcDoRun instead
        # This allows the Windows Service framework to be fully initialized first
        
    def init_hardware_monitor(self):
        """Initialize LibreHardwareMonitor"""
        try:
            servicemanager.LogInfoMsg("=== INITIALIZING LIBREHARDWAREMONITOR ===")
            
            # Get the directory where the service is running
            # Handle both PyInstaller executable and development environments
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller executable
                service_dir = os.path.dirname(sys.executable)
                servicemanager.LogInfoMsg(f"Running as PyInstaller executable")
                servicemanager.LogInfoMsg(f"Executable path: {sys.executable}")
                servicemanager.LogInfoMsg(f"Service directory: {service_dir}")
            else:
                # Running as Python script
                service_dir = os.path.dirname(os.path.abspath(__file__))
                servicemanager.LogInfoMsg(f"Running as Python script")
                servicemanager.LogInfoMsg(f"Script path: {__file__}")
                servicemanager.LogInfoMsg(f"Service directory: {service_dir}")
            
            dll_path = os.path.join(service_dir, "LibreHardwareMonitorLib.dll")
            servicemanager.LogInfoMsg(f"Looking for DLL at: {dll_path}")
            
            if not os.path.exists(dll_path):
                servicemanager.LogErrorMsg(f"LibreHardwareMonitorLib.dll not found at {dll_path}")
                # List files in the service directory for debugging
                try:
                    files_in_dir = os.listdir(service_dir)
                    servicemanager.LogInfoMsg(f"Files in service directory:")
                    for file in sorted(files_in_dir):
                        servicemanager.LogInfoMsg(f"  - {file}")
                except Exception as e:
                    servicemanager.LogErrorMsg(f"Could not list service directory: {e}")
                return
                
            servicemanager.LogInfoMsg(f"Found LibreHardwareMonitorLib.dll")
            servicemanager.LogInfoMsg(f"Loading DLL from: {dll_path}")
            
            clr.AddReference(dll_path)
            from LibreHardwareMonitor import Hardware
            servicemanager.LogInfoMsg(f"Successfully loaded LibreHardwareMonitor assembly")
            
            # Store Hardware module reference for use in other methods
            self.Hardware = Hardware
            
            # Initialize LHM with comprehensive settings
            servicemanager.LogInfoMsg(f"Initializing Hardware.Computer()...")
            self.computer = Hardware.Computer()
            self.computer.IsCpuEnabled = True
            self.computer.IsGpuEnabled = True
            self.computer.IsMemoryEnabled = True
            self.computer.IsMotherboardEnabled = True
            self.computer.IsControllerEnabled = False  # Disable controller sensors (requires HidSharp.dll)
            self.computer.IsStorageEnabled = True      # Enable storage sensors
            
            servicemanager.LogInfoMsg(f"Opening hardware monitoring...")
            self.computer.Open()
            servicemanager.LogInfoMsg(f"Hardware monitoring opened successfully")
            
            # Force an initial update to populate sensors
            servicemanager.LogInfoMsg(f"Performing initial sensor update...")
            for hw in self.computer.Hardware:
                hw.Update()
                for subhw in hw.SubHardware:
                    subhw.Update()
            
            servicemanager.LogInfoMsg("LibreHardwareMonitor initialized successfully")
            servicemanager.LogInfoMsg("=== LIBREHARDWAREMONITOR INITIALIZATION COMPLETE ===")
            
        except Exception as e:
            servicemanager.LogErrorMsg(f"Failed to initialize LibreHardwareMonitor: {e}")
            import traceback
            servicemanager.LogErrorMsg(f"LHM init error traceback: {traceback.format_exc()}")
            self.computer = None
            servicemanager.LogInfoMsg("=== LIBREHARDWAREMONITOR INITIALIZATION FAILED ===")

    def get_cpu_temperature(self):
        """Get CPU temperature using real hardware monitoring."""
        if not self.computer or not self.Hardware:
            servicemanager.LogErrorMsg("LibreHardwareMonitor not initialized - cannot get real temperature")
            return 0  # Return 0 to indicate failure, not fake data
            
        temp = None
        package_temp = None
        core_temp = None
        all_temps = []

        try:
            # Use stored Hardware module reference
            Hardware = self.Hardware
            
            # Update all hardware first
            for hw in self.computer.Hardware:
                hw.Update()
                # Also update sub-hardware (important for some motherboards)
                for subhw in hw.SubHardware:
                    subhw.Update()
            
            for hw in self.computer.Hardware:
                if hw.HardwareType == Hardware.HardwareType.Cpu:
                    servicemanager.LogInfoMsg(f"Found CPU hardware: {hw.Name}")
                    
                    # Check main CPU sensors
                    for sensor in hw.Sensors:
                        if sensor.SensorType == Hardware.SensorType.Temperature:
                            if sensor.Value is not None:
                                sensor_temp = float(sensor.Value)
                                sensor_name = sensor.Name
                                all_temps.append(f"{sensor_name}: {sensor_temp}°C")
                                
                                servicemanager.LogInfoMsg(f"Temperature sensor: {sensor_name} = {sensor_temp}°C")
                                
                                # Prioritize "Core (Tctl/Tdie)" sensor specifically (most accurate for AMD Ryzen)
                                if ("tctl" in sensor_name.lower() or "tdie" in sensor_name.lower()) and "core" in sensor_name.lower():
                                    package_temp = sensor_temp  # This is our preferred sensor
                                    servicemanager.LogInfoMsg(f"Found preferred Core (Tctl/Tdie) temperature: {sensor_temp}°C")
                                # Then any other Tctl/Tdie temperature
                                elif "tctl" in sensor_name.lower() or "tdie" in sensor_name.lower():
                                    if package_temp is None:  # Only use if we don't have Core (Tctl/Tdie)
                                        package_temp = sensor_temp
                                        servicemanager.LogInfoMsg(f"Found AMD Tctl/Tdie temperature: {sensor_temp}°C")
                                # Then Package temperature
                                elif "package" in sensor_name.lower() and sensor_temp > 0:
                                    if package_temp is None:  # Only use if we don't have Tctl/Tdie
                                        package_temp = sensor_temp
                                        servicemanager.LogInfoMsg(f"Found Package temperature: {sensor_temp}°C")
                                # Then other Core temperatures (use average or first valid core)
                                elif "core" in sensor_name.lower() and sensor_temp > 0:
                                    if core_temp is None:
                                        core_temp = sensor_temp
                                    else:
                                        # Average with existing core temp
                                        core_temp = (core_temp + sensor_temp) / 2
                                    servicemanager.LogInfoMsg(f"Found Core temperature: {sensor_temp}°C (avg: {core_temp}°C)")
                                # Any other temperature as fallback
                                elif sensor_temp > 0 and temp is None:
                                    temp = sensor_temp
                                    servicemanager.LogInfoMsg(f"Found other temperature: {sensor_temp}°C")
                    
                    # Check sub-hardware sensors (some CPUs report temps here)
                    for subhw in hw.SubHardware:
                        servicemanager.LogInfoMsg(f"Checking sub-hardware: {subhw.Name}")
                        for sensor in subhw.Sensors:
                            if sensor.SensorType == Hardware.SensorType.Temperature:
                                if sensor.Value is not None:
                                    sensor_temp = float(sensor.Value)
                                    sensor_name = f"{subhw.Name} - {sensor.Name}"
                                    all_temps.append(f"{sensor_name}: {sensor_temp}°C")
                                    
                                    servicemanager.LogInfoMsg(f"Sub-hardware temperature sensor: {sensor_name} = {sensor_temp}°C")
                                    
                                    if "package" in sensor.Name.lower() and sensor_temp > 0:
                                        package_temp = sensor_temp
                                    elif "core" in sensor.Name.lower() and sensor_temp > 0:
                                        if core_temp is None:
                                            core_temp = sensor_temp
                                        else:
                                            core_temp = (core_temp + sensor_temp) / 2
                                    elif sensor_temp > 0 and temp is None:
                                        temp = sensor_temp

                # Also check motherboard sensors for CPU temperature
                elif hw.HardwareType == Hardware.HardwareType.Motherboard:
                    servicemanager.LogInfoMsg(f"Checking motherboard hardware: {hw.Name}")
                    for sensor in hw.Sensors:
                        if sensor.SensorType == Hardware.SensorType.Temperature:
                            if sensor.Value is not None:
                                sensor_temp = float(sensor.Value)
                                sensor_name = f"MB - {sensor.Name}"
                                
                                # Look for CPU-related temperature sensors on motherboard
                                if any(keyword in sensor.Name.lower() for keyword in ["cpu", "processor", "core"]):
                                    all_temps.append(f"{sensor_name}: {sensor_temp}°C")
                                    servicemanager.LogInfoMsg(f"Motherboard CPU temperature sensor: {sensor_name} = {sensor_temp}°C")
                                    
                                    if sensor_temp > 0 and temp is None:
                                        temp = sensor_temp
                    
                    # Check motherboard sub-hardware
                    for subhw in hw.SubHardware:
                        servicemanager.LogInfoMsg(f"Checking motherboard sub-hardware: {subhw.Name}")
                        for sensor in subhw.Sensors:
                            if sensor.SensorType == Hardware.SensorType.Temperature:
                                if sensor.Value is not None:
                                    sensor_temp = float(sensor.Value)
                                    sensor_name = f"MB-{subhw.Name} - {sensor.Name}"
                                    
                                    if any(keyword in sensor.Name.lower() for keyword in ["cpu", "processor", "core"]):
                                        all_temps.append(f"{sensor_name}: {sensor_temp}°C")
                                        servicemanager.LogInfoMsg(f"Motherboard sub-hardware CPU temperature: {sensor_name} = {sensor_temp}°C")
                                        
                                        if sensor_temp > 0 and temp is None:
                                            temp = sensor_temp

            # Determine the best temperature to use
            if package_temp is not None:
                temp = package_temp
                servicemanager.LogInfoMsg(f"Using Package temperature: {temp}°C")
            elif core_temp is not None:
                temp = core_temp
                servicemanager.LogInfoMsg(f"Using Core temperature: {temp}°C")
                        
            if all_temps:
                servicemanager.LogInfoMsg(f"All temperature sensors found: {', '.join(all_temps)}")
            else:
                servicemanager.LogErrorMsg("No temperature sensors found!")
                        
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error reading CPU temperature: {e}")
            import traceback
            servicemanager.LogErrorMsg(f"Temperature error traceback: {traceback.format_exc()}")
            temp = None

        # Return the best temperature we found as integer
        if temp is not None and temp > 0:
            servicemanager.LogInfoMsg(f"Returning real temperature: {int(round(temp))}°C")
            return int(round(temp))
        else:
            servicemanager.LogErrorMsg("Failed to get real CPU temperature - returning 0")
            return 0  # Return 0 to indicate failure, not fake data

    def get_cpu_load(self):
        """Return CPU usage percentage using standard psutil"""
        return self.cpu_monitor.get_cpu_usage()

    def get_master_volume(self):
        """Get master volume level"""
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            return int(volume.GetMasterVolumeLevelScalar() * 100)
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error reading volume: {e}")
            return 0

    def collect_hardware_data(self):
        """Collect all hardware data with consistent format"""
        # Use consistent time format (HH:MM) instead of timestamp
        current_time = datetime.now().strftime("%H:%M")
        
        return {
            "time": current_time,  # Use 'time' format consistently
            "cpu_load": self.get_cpu_load(),
            "volume": self.get_master_volume(),
            "cpu_temp": self.get_cpu_temperature()
        }

    def data_collection_thread(self):
        """Thread that continuously collects hardware data"""
        servicemanager.LogInfoMsg("Data collection thread started")
        
        # Prime psutil CPU counter
        psutil.cpu_percent(interval=None)
        
        while not self.stop_event.is_set():
            try:
                data = self.collect_hardware_data()
                with self.data_lock:
                    self.current_data = data
                    
            except Exception as e:
                servicemanager.LogErrorMsg(f"Error in data collection: {e}")
                
            # Wait 1 second or until stop event
            self.stop_event.wait(1.0)
            
        servicemanager.LogInfoMsg("Data collection thread stopped")

    def named_pipe_server_thread(self):
        """Thread that serves data via named pipe"""
        servicemanager.LogInfoMsg("Named pipe server thread started")
        
        while not self.stop_event.is_set():
            try:
                # Create named pipe
                pipe = win32pipe.CreateNamedPipe(
                    self.pipe_name,
                    win32pipe.PIPE_ACCESS_OUTBOUND,
                    win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
                    1,  # Max instances
                    65536,  # Out buffer size
                    65536,  # In buffer size
                    0,  # Default timeout
                    None  # Security attributes
                )
                
                if pipe == win32file.INVALID_HANDLE_VALUE:
                    servicemanager.LogErrorMsg("Failed to create named pipe")
                    time.sleep(5)
                    continue
                
                servicemanager.LogInfoMsg("Named pipe created, waiting for client connection")
                
                # Wait for client connection
                win32pipe.ConnectNamedPipe(pipe, None)
                servicemanager.LogInfoMsg("Client connected to named pipe")
                
                # Serve data to connected client
                while not self.stop_event.is_set():
                    try:
                        with self.data_lock:
                            data = self.current_data.copy()
                        
                        if data:
                            json_data = json.dumps(data) + '\n'
                            win32file.WriteFile(pipe, json_data.encode('utf-8'))
                        
                        # Send data every second
                        if self.stop_event.wait(1.0):
                            break
                            
                    except pywintypes.error as e:
                        if e.winerror == 232:  # Broken pipe - client disconnected
                            servicemanager.LogInfoMsg("Client disconnected from named pipe")
                            break
                        else:
                            servicemanager.LogErrorMsg(f"Named pipe error: {e}")
                            break
                    except Exception as e:
                        servicemanager.LogErrorMsg(f"Error serving data: {e}")
                        break
                
                # Close pipe
                win32file.CloseHandle(pipe)
                
            except Exception as e:
                servicemanager.LogErrorMsg(f"Named pipe server error: {e}")
                time.sleep(5)
                
        servicemanager.LogInfoMsg("Named pipe server thread stopped")

    def SvcStop(self):
        """Stop the service"""
        servicemanager.LogInfoMsg("Service stop requested")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        
        # Signal threads to stop
        self.stop_event.set()
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """Main service execution"""
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        
        try:
            # Initialize LibreHardwareMonitor after service framework is fully started
            servicemanager.LogInfoMsg("=== INITIALIZING HARDWARE MONITOR SERVICE ===")
            
            # Check administrator privileges
            try:
                is_admin = ctypes.windll.shell32.IsUserAnAdmin()
                servicemanager.LogInfoMsg(f"Administrator privileges: {'YES' if is_admin else 'NO'}")
                if not is_admin:
                    servicemanager.LogErrorMsg("WARNING: Not running as administrator! Temperature sensors may not work.")
                else:
                    servicemanager.LogInfoMsg("Running with administrator privileges")
            except:
                servicemanager.LogInfoMsg("Could not check administrator status")
            
            # Initialize LibreHardwareMonitor now that service framework is ready
            servicemanager.LogInfoMsg("Initializing LibreHardwareMonitor...")
            self.init_hardware_monitor()
            
            if not self.computer:
                servicemanager.LogErrorMsg("LibreHardwareMonitor initialization failed - service cannot continue")
                return
            
            servicemanager.LogInfoMsg("LibreHardwareMonitor initialized successfully")
            
            # Initialize psutil CPU monitoring
            servicemanager.LogInfoMsg("Initializing psutil CPU monitoring...")
            if not self.cpu_monitor.initialize():
                servicemanager.LogErrorMsg("psutil CPU monitoring initialization failed - will use basic fallback")
            else:
                servicemanager.LogInfoMsg("psutil CPU monitoring initialized successfully")
            
            # Start data collection thread
            data_thread = threading.Thread(target=self.data_collection_thread, daemon=True)
            data_thread.start()
            
            # Start named pipe server thread
            pipe_thread = threading.Thread(target=self.named_pipe_server_thread, daemon=True)
            pipe_thread.start()
            
            servicemanager.LogInfoMsg("Hardware Monitor Service started successfully")
            servicemanager.LogInfoMsg("Serving data via named pipe...")
            
            # Wait for stop event
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
            # Wait for threads to finish
            servicemanager.LogInfoMsg("Waiting for threads to finish...")
            data_thread.join(timeout=5)
            pipe_thread.join(timeout=5)
            
            # Clean up resources
            servicemanager.LogInfoMsg("Cleaning up resources...")
            self.cpu_monitor.cleanup()
            
            if self.computer:
                try:
                    self.computer.Close()
                    servicemanager.LogInfoMsg("LibreHardwareMonitor closed successfully")
                except Exception as e:
                    servicemanager.LogErrorMsg(f"Error closing LibreHardwareMonitor: {e}")
            
        except Exception as e:
            servicemanager.LogErrorMsg(f"Service error: {e}")
            import traceback
            servicemanager.LogErrorMsg(f"Service error traceback: {traceback.format_exc()}")
            
            # Clean up resources even on error
            try:
                self.cpu_monitor.cleanup()
                if self.computer:
                    self.computer.Close()
            except:
                pass
        
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STOPPED,
            (self._svc_name_, '')
        )


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(HardwareMonitorService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(HardwareMonitorService)

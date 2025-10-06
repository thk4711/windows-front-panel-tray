"""
Tray Serial Monitor Client - GUI Component
This client runs in user space and connects to the hardware monitoring service.
It handles the system tray interface and serial communication.
"""

import json
import threading
import time
import sys
import os
from datetime import datetime
import win32pipe
import win32file
import pywintypes

import serial
from PIL import Image, ImageDraw
import pystray

# Import our ESP32 port detector
from esp32_port_detector import ESP32PortDetector

# ===== CONFIGURATION =====
# No config file needed - using hardcoded defaults
BAUD_RATE = 115200


class HardwareDataClient:
    """Client that connects to the hardware monitoring service via named pipe"""

    def __init__(self):
        self.pipe_name = r'\\.\pipe\TrayHardwareMonitor'
        self.pipe = None
        self.connected = False
        self.last_data = {}
        self.data_lock = threading.Lock()
        self.stop_event = threading.Event()

    def connect_to_service(self):
        """Connect to the hardware monitoring service"""
        try:
            # Try to connect to the named pipe
            self.pipe = win32file.CreateFile(
                self.pipe_name,
                win32file.GENERIC_READ,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            )

            if self.pipe != win32file.INVALID_HANDLE_VALUE:
                self.connected = True
                print("Connected to hardware monitoring service")
                return True
            else:
                print("Failed to connect to hardware monitoring service")
                return False

        except pywintypes.error as e:
            print(f"Error connecting to service: {e}")
            return False

    def disconnect_from_service(self):
        """Disconnect from the service"""
        if self.pipe and self.pipe != win32file.INVALID_HANDLE_VALUE:
            win32file.CloseHandle(self.pipe)
            self.pipe = None
        self.connected = False
        print("Disconnected from hardware monitoring service")

    def read_data_thread(self):
        """Thread that continuously reads data from the service"""
        print("Hardware data reader thread started")

        while not self.stop_event.is_set():
            if not self.connected:
                # Try to connect
                if self.connect_to_service():
                    continue
                else:
                    # Wait before retrying
                    if self.stop_event.wait(5.0):
                        break
                    continue

            try:
                # Read data from pipe
                result, data = win32file.ReadFile(self.pipe, 4096)
                if result == 0 and data:
                    # Parse JSON data
                    json_str = data.decode('utf-8').strip()
                    if json_str:
                        print(json_str)
                        hardware_data = json.loads(json_str)
                        with self.data_lock:
                            self.last_data = hardware_data

            except pywintypes.error as e:
                if e.winerror == 109:  # Broken pipe
                    print("Service disconnected, will retry...")
                    self.disconnect_from_service()
                else:
                    print(f"Pipe read error: {e}")
                    self.disconnect_from_service()
            except json.JSONDecodeError as e:
                print(f"JSON decode error: {e}")
            except Exception as e:
                print(f"Unexpected error reading from service: {e}")
                self.disconnect_from_service()

        self.disconnect_from_service()
        print("Hardware data reader thread stopped")

    def get_hardware_data(self):
        """Get the latest hardware data from the service"""
        with self.data_lock:
            return self.last_data.copy()

    def start(self):
        """Start the data reading thread"""
        self.reader_thread = threading.Thread(target=self.read_data_thread, daemon=True)
        self.reader_thread.start()

    def stop(self):
        """Stop the data reading thread"""
        self.stop_event.set()
        self.disconnect_from_service()


def get_time_str():
    return datetime.now().strftime("%H:%M")

def collect_data(hardware_client):
    """Collect data combining service data with local data"""
    hardware_data = hardware_client.get_hardware_data()

    # Use hardware data from service, fallback to defaults if not available
    return {
        "time": get_time_str(),
        "cpu_load": hardware_data.get("cpu_load", 0),
        "volume": hardware_data.get("volume", 0),
        "cpu_temp": hardware_data.get("cpu_temp", 0)
    }

# ===== Enhanced Serial worker thread with auto-detection =====
def serial_worker(stop_event, hardware_client):
    """Enhanced serial worker with ESP32 auto-detection and reconnection"""
    detector = ESP32PortDetector()
    ser = None
    current_port = None
    last_port_scan = 0
    port_scan_interval = 10  # Scan for new ports every 10 seconds
    
    def cleanup_serial():
        nonlocal ser
        if ser and ser.is_open:
            try:
                ser.close()
                print(f"Closed serial connection on {current_port}")
            except:
                pass
        ser = None
    
    def try_connect_to_port(port):
        """Try to connect to a specific port"""
        try:
            test_ser = serial.Serial(port, BAUD_RATE, timeout=1)
            print(f"Serial connection opened on {port} at {BAUD_RATE} baud.")
            return test_ser
        except serial.SerialException as e:
            print(f"Failed to connect to {port}: {e}")
            return None
    
    print("Starting enhanced serial worker with ESP32 auto-detection...")
    
    while not stop_event.is_set():
        current_time = time.time()
        
        # Check if we need to scan for ports (either no connection or periodic scan)
        if (ser is None or not ser.is_open or 
            current_time - last_port_scan > port_scan_interval):
            
            print("Scanning for ESP32 devices...")
            
            # Try auto-detection
            detected_port = detector.get_best_esp32_port(test_connection=False)
            
            if detected_port:
                print(f"Auto-detected ESP32 on {detected_port}")
                
                # If this is a new port, try to connect
                if detected_port != current_port:
                    cleanup_serial()
                    ser = try_connect_to_port(detected_port)
                    if ser:
                        current_port = detected_port
            else:
                print("No ESP32 devices found")
                cleanup_serial()
                current_port = None
            
            last_port_scan = current_time
        
        # Try to send data if we have a connection
        if ser and ser.is_open:
            payload = collect_data(hardware_client)
            try:
                ser.write(json.dumps(payload).encode() + b"\n")
            except serial.SerialException as e:
                print(f"Serial communication error on {current_port}: {e}")
                print("Will attempt to reconnect...")
                cleanup_serial()
                current_port = None
            except Exception as e:
                print(f"Unexpected error sending data: {e}")
        else:
            # No connection available, wait a bit before trying again
            if current_time - last_port_scan > 5:  # Only print this message occasionally
                print("No ESP32 connection available, scanning...")
        
        # Wait before next iteration
        time.sleep(1)
    
    # Cleanup on exit
    cleanup_serial()
    print("Serial worker thread stopped")

# ===== Tray icon =====
def create_image():
    """Create a tray icon representing hardware monitoring and serial communication"""
    img = Image.new('RGBA', (64, 64), color=(0, 0, 0, 0))  # Transparent background
    d = ImageDraw.Draw(img)
    
    # Draw a computer/CPU chip representation - using more of the available space
    # Main chip body (dark blue) - expanded to use more of the canvas
    d.rectangle((4, 4, 60, 60), fill=(30, 60, 120), outline=(60, 100, 180), width=2)
    
    # CPU pins/connections (silver/gray) - extended to edges
    # Top pins
    for x in range(8, 57, 8):
        d.rectangle((x, 0, x+3, 4), fill=(180, 180, 180))
    # Bottom pins  
    for x in range(8, 57, 8):
        d.rectangle((x, 60, x+3, 64), fill=(180, 180, 180))
    # Left pins
    for y in range(8, 57, 8):
        d.rectangle((0, y, 4, y+3), fill=(180, 180, 180))
    # Right pins
    for y in range(8, 57, 8):
        d.rectangle((60, y, 64, y+3), fill=(180, 180, 180))
    
    # Draw signal waves (representing serial communication) - larger and more prominent
    # Green signal lines
    signal_color = (0, 220, 0)
    d.arc((8, 8, 24, 24), 0, 180, fill=signal_color, width=3)
    d.arc((12, 12, 28, 28), 0, 180, fill=signal_color, width=3)
    d.arc((16, 16, 32, 32), 0, 180, fill=signal_color, width=3)
    
    # Add a temperature indicator (red dot) - larger and more visible
    d.ellipse((44, 12, 52, 20), fill=(255, 60, 60))
    
    return img

def create_menu(hardware_client):
    """Create tray menu"""
    def on_exit(icon, item):
        stop_event.set()
        icon.stop()

    return pystray.Menu(
        pystray.MenuItem("Exit", on_exit)
    )

def update_tooltip(icon, hardware_client):
    """Update the tray icon tooltip with connection status"""
    if hardware_client.connected:
        tooltip = "Tray Serial Monitor - Service Connected"
    else:
        tooltip = "Tray Serial Monitor - Service Disconnected"
    
    icon.title = tooltip

def tooltip_updater(icon, hardware_client, stop_event):
    """Thread to periodically update the tooltip"""
    while not stop_event.is_set():
        update_tooltip(icon, hardware_client)
        if stop_event.wait(2.0):  # Update every 2 seconds
            break

if __name__ == "__main__":
    print("Starting Tray Serial Monitor Client...")

    # Create hardware data client
    hardware_client = HardwareDataClient()
    hardware_client.start()

    # Give the client a moment to connect
    time.sleep(2)

    # Test data collection
    print("Testing data collection...")
    try:
        test_data = collect_data(hardware_client)
        print(f"Test data: {test_data}")
    except Exception as e:
        print(f"Error in data collection: {e}")

    # Global stop event
    stop_event = threading.Event()

    # Start serial sending thread
    serial_thread = threading.Thread(target=serial_worker, args=(stop_event, hardware_client), daemon=True)
    serial_thread.start()
    print("Serial worker thread started")

    # Create system tray icon
    print("Creating system tray icon...")
    try:
        icon = pystray.Icon(
            "TraySerialClient",
            create_image(),
            menu=create_menu(hardware_client),
            title="Tray Serial Monitor - Starting..."
        )
        
        # Start tooltip updater thread
        tooltip_thread = threading.Thread(target=tooltip_updater, args=(icon, hardware_client, stop_event), daemon=True)
        tooltip_thread.start()
        print("Tooltip updater thread started")
        
        print("Running system tray...")
        icon.run()
    except Exception as e:
        print(f"Error creating system tray icon: {e}")
        print("Running without system tray...")
        # Keep the script running without tray icon
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("Stopping...")
            stop_event.set()

    # Cleanup
    print("Shutting down...")
    stop_event.set()
    hardware_client.stop()

    # Wait for threads to finish
    serial_thread.join(timeout=5)
    print("Client shutdown complete")

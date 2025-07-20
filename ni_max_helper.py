import os
import sys
import subprocess
import pyvisa
import time
import re
import winreg
import glob

class NIMaxHelper:
    def __init__(self):
        self.ni_max_path = r"C:\Program Files (x86)\National Instruments\MAX\NIMax.exe"
        self.rm = None

    def check_gpib_drivers(self):
        """Check for GPIB drivers in all possible locations"""
        print("\n=== Checking GPIB Drivers ===")
        
        # Additional GPIB driver locations
        gpib_paths = [
            r"C:\Program Files (x86)\National Instruments\NI-488.2",
            r"C:\Program Files (x86)\National Instruments\Shared\GPIB",
            r"C:\Program Files\National Instruments\NI-488.2",
            r"C:\Program Files\National Instruments\Shared\GPIB",
            r"C:\Windows\System32",
            r"C:\Windows\SysWOW64"
        ]

        for base_path in gpib_paths:
            if os.path.exists(base_path):
                print(f"\nChecking in: {base_path}")
                # Look for common GPIB-related files
                gpib_files = [
                    'gpib-32.dll',
                    'Gpib-32.dll',
                    'ni4882.dll',
                    'nigpib32.dll',
                    'nigpibc.dll'
                ]
                
                found_files = []
                for file in gpib_files:
                    file_path = os.path.join(base_path, file)
                    if os.path.exists(file_path):
                        size = os.path.getsize(file_path)
                        modified = time.strftime('%Y-%m-%d', time.localtime(os.path.getmtime(file_path)))
                        found_files.append(f"  ✓ {file} (Size: {size:,} bytes, Modified: {modified})")
                
                if found_files:
                    print("Found GPIB files:")
                    for file in found_files:
                        print(file)
                else:
                    print("  No GPIB files found in this location")
            else:
                print(f"\n✗ Path not found: {base_path}")

    def check_ni_488_registry(self):
        """Check NI-488.2 specific registry entries"""
        print("\n=== Checking NI-488.2 Registry ===")
        
        registry_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\National Instruments\NI-488.2"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\National Instruments\NI-488.2"),
        ]

        for root, path in registry_paths:
            try:
                key = winreg.OpenKey(root, path, 0, winreg.KEY_READ)
                print(f"\nFound registry key: {path}")
                
                try:
                    i = 0
                    while True:
                        name, value, type_ = winreg.EnumValue(key, i)
                        print(f"  {name} = {value}")
                        i += 1
                except WindowsError:
                    pass
                
                winreg.CloseKey(key)
            except WindowsError:
                print(f"✗ Registry key not found: {path}")

    def check_device_manager_gpib(self):
        """Check Device Manager specifically for GPIB devices"""
        print("\n=== Checking Device Manager for GPIB Hardware ===")
        try:
            # Use Windows Management Instrumentation (WMI) to query devices
            cmd = 'wmic path win32_pnpentity get caption,deviceid,status /format:list'
            output = subprocess.check_output(cmd, shell=True).decode('utf-8', errors='ignore')
            
            # Look for GPIB-related devices
            gpib_devices = []
            current_device = {}
            
            for line in output.split('\n'):
                line = line.strip()
                if line:
                    if '=' in line:
                        key, value = line.split('=', 1)
                        current_device[key.lower()] = value
                else:
                    if current_device and any('GPIB' in v.upper() for v in current_device.values()):
                        gpib_devices.append(current_device)
                    current_device = {}
            
            if gpib_devices:
                print("\nFound GPIB devices:")
                for device in gpib_devices:
                    print(f"\nDevice: {device.get('caption', 'Unknown')}")
                    print(f"ID: {device.get('deviceid', 'Unknown')}")
                    print(f"Status: {device.get('status', 'Unknown')}")
            else:
                print("\nNo GPIB devices found in Device Manager")
                print("\nTroubleshooting steps:")
                print("1. Check if GPIB adapter is physically connected")
                print("2. Try a different USB port")
                print("3. Check if the adapter appears under 'Other devices' or with warning symbol")
                print("4. Try uninstalling and reinstalling NI-488.2")
        except Exception as e:
            print(f"Error checking Device Manager: {e}")

def main():
    helper = NIMaxHelper()
    
    # Check GPIB drivers
    helper.check_gpib_drivers()
    
    # Check NI-488.2 registry entries
    helper.check_ni_488_registry()
    
    # Check Device Manager for GPIB hardware
    helper.check_device_manager_gpib()
    
    print("\n=== Recommendations ===")
    print("1. If no GPIB drivers found:")
    print("   - Repair/reinstall NI-488.2 from NI Package Manager")
    print("   - Download location: ni.com/downloads")
    print("2. If drivers exist but no devices detected:")
    print("   - Check USB-GPIB adapter connection")
    print("   - Try different USB ports")
    print("   - Check Windows Device Manager for issues")
    print("3. After any changes:")
    print("   - Restart the computer")
    print("   - Run this script again to verify")

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
APSM2000 USB Communication Test Script
Tests basic USB connectivity and helps diagnose common issues.
"""
import ctypes
import time
import os
import sys

# Constants
VID = 0x10C4  # 4292 decimal
PID = 0x8805  # 34869 decimal
HID_UART_SUCCESS = 0x00

class APSM2000_Tester:
    def __init__(self, debug=True):
        self.debug = debug
        self.dev_handle = None
        self.dll = None
        
    def log(self, msg):
        if self.debug:
            print(f"[DEBUG] {msg}")
            
    def error(self, msg):
        print(f"[ERROR] {msg}")
        
    def load_dll(self):
        """Step 1: Load the Silicon Labs DLL"""
        try:
            # Try current directory first
            dll_path = os.path.join(os.path.dirname(__file__), "SLABHIDtoUART.dll")
            if not os.path.exists(dll_path):
                self.error(f"DLL not found at: {dll_path}")
                return False
                
            self.dll = ctypes.WinDLL(dll_path)
            self.log("DLL loaded successfully")
            return True
            
        except Exception as e:
            self.error(f"Failed to load DLL: {e}")
            return False
            
    def find_devices(self):
        """Step 2: Look for APSM2000 devices"""
        if not self.dll:
            self.error("DLL not loaded")
            return 0
            
        try:
            num_devices = ctypes.c_ulong(0)
            ret = self.dll.HidUart_GetNumDevices(
                ctypes.byref(num_devices),
                VID,
                PID
            )
            
            if ret != HID_UART_SUCCESS:
                self.error(f"GetNumDevices failed with error {ret}")
                return 0
                
            self.log(f"Found {num_devices.value} device(s)")
            return num_devices.value
            
        except Exception as e:
            self.error(f"Error checking for devices: {e}")
            return 0
            
    def open_device(self, device_index=0):
        """Step 3: Try to open the device"""
        if not self.dll:
            self.error("DLL not loaded")
            return False
            
        try:
            self.dev_handle = ctypes.c_void_p()
            ret = self.dll.HidUart_Open(
                ctypes.byref(self.dev_handle),
                device_index,
                VID,
                PID
            )
            
            if ret != HID_UART_SUCCESS:
                self.error(f"Open failed with error {ret}")
                return False
                
            self.log("Device opened successfully")
            return True
            
        except Exception as e:
            self.error(f"Error opening device: {e}")
            return False
            
    def configure_device(self):
        """Step 4: Configure the UART settings"""
        if not self.dev_handle:
            self.error("Device not open")
            return False
            
        try:
            # Configure for 115200 8N1 with RTS/CTS
            ret = self.dll.HidUart_SetUartConfig(
                self.dev_handle,    # handle
                115200,            # baud
                8,                 # data bits
                0,                 # no parity
                0,                 # 1 stop bit
                2                  # RTS/CTS flow control
            )
            
            if ret != HID_UART_SUCCESS:
                self.error(f"SetUartConfig failed with error {ret}")
                return False
                
            # Set reasonable timeouts (1 second)
            ret = self.dll.HidUart_SetTimeouts(
                self.dev_handle,
                1000,  # read timeout ms
                1000   # write timeout ms
            )
            
            if ret != HID_UART_SUCCESS:
                self.error(f"SetTimeouts failed with error {ret}")
                return False
                
            self.log("Device configured successfully")
            return True
            
        except Exception as e:
            self.error(f"Error configuring device: {e}")
            return False
            
    def test_communication(self):
        """Step 5: Test basic communication"""
        if not self.dev_handle:
            self.error("Device not open")
            return False
            
        try:
            # 1. Clear any pending data
            self.dll.HidUart_FlushBuffers(self.dev_handle, True, True)
            time.sleep(0.1)
            
            # 2. Send *CLS command
            cmd = "*CLS\n"
            cmd_bytes = cmd.encode('ascii')
            written = ctypes.c_ulong(0)
            
            ret = self.dll.HidUart_Write(
                self.dev_handle,
                cmd_bytes,
                len(cmd_bytes),
                ctypes.byref(written)
            )
            
            if ret != HID_UART_SUCCESS:
                self.error(f"Write failed with error {ret}")
                return False
                
            if written.value != len(cmd_bytes):
                self.error(f"Incomplete write: {written.value} of {len(cmd_bytes)}")
                return False
                
            self.log("Basic write test passed")
            
            # 3. Small delay
            time.sleep(0.5)
            
            # 4. Try reading (shouldn't be any data, but checks if read works)
            buffer = (ctypes.c_ubyte * 64)()
            bytes_read = ctypes.c_ulong(0)
            
            ret = self.dll.HidUart_Read(
                self.dev_handle,
                buffer,
                64,
                ctypes.byref(bytes_read)
            )
            
            if ret != HID_UART_SUCCESS:
                self.error(f"Read failed with error {ret}")
                return False
                
            self.log("Basic read test passed")
            return True
            
        except Exception as e:
            self.error(f"Error testing communication: {e}")
            return False
            
    def close(self):
        """Clean up"""
        if self.dev_handle:
            try:
                self.dll.HidUart_Close(self.dev_handle)
                self.log("Device closed")
            except:
                pass
            self.dev_handle = None

def main():
    print("\n=== APSM2000 USB Communication Diagnostic ===\n")
    
    # Create tester
    tester = APSM2000_Tester(debug=True)
    
    try:
        # Step 1: Load DLL
        print("\nStep 1: Loading DLL...")
        if not tester.load_dll():
            print("Failed to load DLL. Cannot continue.")
            return
        print("DLL loaded successfully.")
        
        # Step 2: Find devices
        print("\nStep 2: Looking for APSM2000 devices...")
        num_devices = tester.find_devices()
        if num_devices == 0:
            print("No APSM2000 devices found. Cannot continue.")
            return
        print(f"Found {num_devices} device(s).")
        
        # Step 3: Open device
        print("\nStep 3: Opening device...")
        if not tester.open_device():
            print("Failed to open device. Cannot continue.")
            return
        print("Device opened successfully.")
        
        # Step 4: Configure
        print("\nStep 4: Configuring device...")
        if not tester.configure_device():
            print("Failed to configure device.")
            return
        print("Device configured successfully.")
        
        # Step 5: Test communication
        print("\nStep 5: Testing basic communication...")
        if not tester.test_communication():
            print("Communication test failed.")
            return
        print("Communication test passed.")
        
        print("\nAll tests completed successfully!")
        
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        
    finally:
        # Clean up
        tester.close()
        print("\nTest complete.")

if __name__ == "__main__":
    main()
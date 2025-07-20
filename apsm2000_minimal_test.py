#!/usr/bin/env python3
"""
Minimal APSM2000 Communication Test
Tests only the most basic USB-to-RS232 communication.
No streaming, just simple command/response.
"""
import ctypes
import time
import os
import sys

# Constants
VID = 0x10C4
PID = 0x8805
HID_UART_SUCCESS = 0x00

def print_debug(msg):
    print(f"[DEBUG] {msg}")
    sys.stdout.flush()  # Ensure output is shown immediately

def test_basic_communication():
    print_debug("\n=== APSM2000 Minimal Communication Test ===\n")
    
    try:
        # 1. Load DLL
        print_debug("Loading DLL...")
        dll_path = os.path.join(os.path.dirname(__file__), "SLABHIDtoUART.dll")
        if not os.path.exists(dll_path):
            print_debug(f"Error: DLL not found at {dll_path}")
            return False
            
        dll = ctypes.WinDLL(dll_path)
        print_debug("DLL loaded successfully")
        
        # 2. Check for devices
        print_debug("\nLooking for devices...")
        num_devices = ctypes.c_ulong(0)
        ret = dll.HidUart_GetNumDevices(ctypes.byref(num_devices), VID, PID)
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: GetNumDevices failed with {ret}")
            return False
            
        if num_devices.value == 0:
            print_debug("Error: No devices found")
            return False
            
        print_debug(f"Found {num_devices.value} device(s)")
        
        # 3. Open device
        print_debug("\nOpening device...")
        dev_handle = ctypes.c_void_p()
        ret = dll.HidUart_Open(ctypes.byref(dev_handle), 0, VID, PID)
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: Open failed with {ret}")
            return False
            
        print_debug("Device opened successfully")
        
        # 4. Configure UART
        print_debug("\nConfiguring UART...")
        ret = dll.HidUart_SetUartConfig(
            dev_handle,
            115200,    # baud
            8,         # data bits
            0,         # no parity
            0,         # 1 stop bit
            2          # RTS/CTS flow control
        )
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: SetUartConfig failed with {ret}")
            return False
            
        # 5. Set longer timeouts
        print_debug("Setting timeouts...")
        ret = dll.HidUart_SetTimeouts(dev_handle, 5000, 5000)  # 5 second timeouts
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: SetTimeouts failed with {ret}")
            return False
            
        # 6. Flush buffers
        print_debug("Flushing buffers...")
        ret = dll.HidUart_FlushBuffers(dev_handle, True, True)
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: FlushBuffers failed with {ret}")
            return False
            
        # 7. Wait for device to stabilize
        print_debug("Waiting for device to stabilize...")
        time.sleep(2)
        
        # 8. Try simple write/read test
        print_debug("\nTesting basic communication...")
        
        # Clear interface
        print_debug("Sending *CLS...")
        cmd = "*CLS\n"
        cmd_bytes = cmd.encode('ascii')
        written = ctypes.c_ulong(0)
        
        ret = dll.HidUart_Write(dev_handle, cmd_bytes, len(cmd_bytes), 
                               ctypes.byref(written))
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: Write failed with {ret}")
            return False
            
        if written.value != len(cmd_bytes):
            print_debug(f"Error: Incomplete write {written.value}/{len(cmd_bytes)}")
            return False
            
        print_debug("*CLS sent successfully")
        time.sleep(1)  # Wait after clear
        
        # Try to get ID
        print_debug("\nRequesting device ID...")
        cmd = "*IDN?\n"
        cmd_bytes = cmd.encode('ascii')
        written = ctypes.c_ulong(0)
        
        # Flush before write
        dll.HidUart_FlushBuffers(dev_handle, True, True)
        time.sleep(0.1)
        
        ret = dll.HidUart_Write(dev_handle, cmd_bytes, len(cmd_bytes), 
                               ctypes.byref(written))
        if ret != HID_UART_SUCCESS:
            print_debug(f"Error: Write failed with {ret}")
            return False
            
        print_debug("*IDN? sent successfully")
        
        # Read response with timeout
        print_debug("Waiting for response...")
        start_time = time.time()
        timeout = 5.0  # 5 seconds
        response = bytearray()
        
        while (time.time() - start_time) < timeout:
            buffer = (ctypes.c_ubyte * 64)()
            bytes_read = ctypes.c_ulong(0)
            
            ret = dll.HidUart_Read(dev_handle, buffer, 64, ctypes.byref(bytes_read))
            if ret != HID_UART_SUCCESS:
                print_debug(f"Error: Read failed with {ret}")
                return False
                
            if bytes_read.value > 0:
                response.extend(buffer[:bytes_read.value])
                if b'\n' in response:
                    break
                    
            time.sleep(0.1)  # Small delay between reads
            
        if b'\n' not in response:
            print_debug("Error: No response received within timeout")
            return False
            
        # Convert response to string
        response_str = response.split(b'\n')[0].decode('ascii').strip()
        print_debug(f"\nDevice ID Response: {response_str}")
        
        # Close device
        print_debug("\nClosing device...")
        dll.HidUart_Close(dev_handle)
        print_debug("Device closed successfully")
        
        return True
        
    except Exception as e:
        print_debug(f"\nUnexpected error: {e}")
        return False
        
if __name__ == "__main__":
    # Run test
    if test_basic_communication():
        print_debug("\nBasic communication test PASSED!")
    else:
        print_debug("\nBasic communication test FAILED!")
    
    print_debug("\n=== Test Complete ===\n")
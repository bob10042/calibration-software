#!/usr/bin/env python3
"""
APSM2000 Basic Communication Test
Specifically designed for USB-to-RS232 adapter scenarios.
Features:
- Minimal commands to test basic connectivity
- Robust error handling
- Slower timing to accommodate adapter latency
- Debug logging
"""

import ctypes
import time
import os
import sys
import logging
from datetime import datetime

# Setup logging
log_filename = f"apsm2000_debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler(sys.stdout)
    ]
)

# Constants
VID = 0x10C4  # 4292 decimal
PID = 0x8805  # 34869 decimal
HID_UART_SUCCESS = 0x00

class APSM2000_Basic:
    def __init__(self):
        self.dll = None
        self.dev_handle = None
        self.connected = False
        
    def load_dll(self):
        """Load the Silicon Labs DLL with error checking"""
        try:
            dll_path = os.path.join(os.path.dirname(__file__), "SLABHIDtoUART.dll")
            if not os.path.exists(dll_path):
                logging.error(f"DLL not found at: {dll_path}")
                return False
                
            self.dll = ctypes.WinDLL(dll_path)
            logging.info("DLL loaded successfully")
            return True
            
        except Exception as e:
            logging.error(f"Failed to load DLL: {e}")
            return False
            
    def open_device(self):
        """Open and configure the device with USB-to-RS232 considerations"""
        try:
            # 1. Check for devices
            num_devices = ctypes.c_ulong(0)
            ret = self.dll.HidUart_GetNumDevices(
                ctypes.byref(num_devices),
                VID,
                PID
            )
            
            if ret != HID_UART_SUCCESS:
                logging.error(f"GetNumDevices failed: {ret}")
                return False
                
            if num_devices.value == 0:
                logging.error("No APSM2000 devices found")
                return False
                
            logging.info(f"Found {num_devices.value} device(s)")
            
            # 2. Open first device
            self.dev_handle = ctypes.c_void_p()
            ret = self.dll.HidUart_Open(
                ctypes.byref(self.dev_handle),
                0,  # First device
                VID,
                PID
            )
            
            if ret != HID_UART_SUCCESS:
                logging.error(f"Open failed: {ret}")
                return False
                
            # 3. Configure UART - Important for USB-to-RS232
            ret = self.dll.HidUart_SetUartConfig(
                self.dev_handle,
                115200,        # Baud rate
                8,             # Data bits
                0,             # No parity
                0,             # 1 stop bit
                2             # RTS/CTS flow control
            )
            
            if ret != HID_UART_SUCCESS:
                logging.error(f"SetUartConfig failed: {ret}")
                return False
                
            # 4. Longer timeouts for USB-to-RS232
            ret = self.dll.HidUart_SetTimeouts(
                self.dev_handle,
                2000,  # 2 second read timeout
                2000   # 2 second write timeout
            )
            
            if ret != HID_UART_SUCCESS:
                logging.error(f"SetTimeouts failed: {ret}")
                return False
                
            # 5. Flush buffers
            ret = self.dll.HidUart_FlushBuffers(
                self.dev_handle,
                True,  # Flush transmit
                True   # Flush receive
            )
            
            if ret != HID_UART_SUCCESS:
                logging.error(f"FlushBuffers failed: {ret}")
                return False
                
            # 6. Wait for adapter to stabilize
            time.sleep(2)
            
            self.connected = True
            logging.info("Device opened and configured successfully")
            return True
            
        except Exception as e:
            logging.error(f"Error in open_device: {e}")
            return False
            
    def write_command(self, command, max_retries=3):
        """Write command with retry logic and chunking for USB-to-RS232"""
        if not self.connected:
            logging.error("Device not connected")
            return False
            
        # Add newline if not present
        if not command.endswith('\n'):
            command += '\n'
            
        cmd_bytes = command.encode('ascii')
        chunk_size = 32  # Smaller chunks for USB-to-RS232
        
        for attempt in range(max_retries):
            try:
                # Flush before write
                self.dll.HidUart_FlushBuffers(self.dev_handle, True, True)
                time.sleep(0.1)  # Short delay after flush
                
                # Write in chunks
                for i in range(0, len(cmd_bytes), chunk_size):
                    chunk = cmd_bytes[i:i+chunk_size]
                    written = ctypes.c_ulong(0)
                    
                    ret = self.dll.HidUart_Write(
                        self.dev_handle,
                        chunk,
                        len(chunk),
                        ctypes.byref(written)
                    )
                    
                    if ret != HID_UART_SUCCESS:
                        raise IOError(f"Write failed: {ret}")
                        
                    if written.value != len(chunk):
                        raise IOError(f"Incomplete write: {written.value}/{len(chunk)}")
                        
                    time.sleep(0.05)  # Delay between chunks
                    
                logging.debug(f"Successfully wrote command: {command.strip()}")
                return True
                
            except Exception as e:
                logging.warning(f"Write attempt {attempt + 1} failed: {e}")
                time.sleep(0.5)  # Wait before retry
                
        logging.error(f"All write attempts failed for command: {command.strip()}")
        return False
        
    def read_response(self, timeout_sec=2.0):
        """Read response with timeout and validation"""
        if not self.connected:
            logging.error("Device not connected")
            return None
            
        try:
            start_time = time.time()
            response = bytearray()
            
            while (time.time() - start_time) < timeout_sec:
                # Read in chunks
                buffer = (ctypes.c_ubyte * 64)()
                bytes_read = ctypes.c_ulong(0)
                
                ret = self.dll.HidUart_Read(
                    self.dev_handle,
                    buffer,
                    64,
                    ctypes.byref(bytes_read)
                )
                
                if ret != HID_UART_SUCCESS:
                    logging.error(f"Read failed: {ret}")
                    return None
                    
                if bytes_read.value > 0:
                    response.extend(buffer[:bytes_read.value])
                    if b'\n' in response:
                        break
                        
                time.sleep(0.05)  # Small delay between reads
                
            if b'\n' not in response:
                logging.error("Timeout waiting for response")
                return None
                
            # Convert to string and clean up
            result = response.split(b'\n')[0].decode('ascii').strip()
            logging.debug(f"Read response: {result}")
            return result
            
        except Exception as e:
            logging.error(f"Error reading response: {e}")
            return None
            
    def close(self):
        """Clean up resources"""
        if self.dev_handle:
            try:
                self.dll.HidUart_Close(self.dev_handle)
                logging.info("Device closed")
            except:
                pass
            self.dev_handle = None
            self.connected = False

def main():
    """
    Basic communication test sequence
    """
    logging.info("=== Starting APSM2000 Basic Communication Test ===")
    
    m2000 = APSM2000_Basic()
    
    try:
        # 1. Load DLL
        if not m2000.load_dll():
            logging.error("Failed to load DLL")
            return
            
        # 2. Open device
        if not m2000.open_device():
            logging.error("Failed to open device")
            return
            
        # 3. Clear interface
        logging.info("Sending *CLS...")
        if not m2000.write_command("*CLS"):
            logging.error("Failed to send *CLS")
            return
            
        time.sleep(1)  # Wait after clear
        
        # 4. Try to get ID
        logging.info("Requesting device ID...")
        if not m2000.write_command("*IDN?"):
            logging.error("Failed to send *IDN?")
            return
            
        response = m2000.read_response(timeout_sec=3.0)
        if response:
            logging.info(f"Device ID: {response}")
        else:
            logging.error("Failed to get device ID")
            return
            
        # 5. Try single voltage read
        logging.info("Reading CH1 voltage...")
        if not m2000.write_command("READ? VOLTS:CH1:ACDC"):
            logging.error("Failed to send voltage query")
            return
            
        response = m2000.read_response(timeout_sec=3.0)
        if response:
            try:
                voltage = float(response)
                logging.info(f"CH1 Voltage: {voltage:.3f}V")
            except ValueError:
                logging.error(f"Invalid voltage response: {response}")
        else:
            logging.error("Failed to read voltage")
            
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        
    finally:
        m2000.close()
        logging.info("=== Test Complete ===")

if __name__ == "__main__":
    main()
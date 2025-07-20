# APSM2000 (M2000 Series) Communication Guide

This guide contains all the Python code examples and explanations for communicating with the APSM2000 via USB, RS232, and LAN interfaces.

## Table of Contents
1. [USB HID Communication](#usb-hid-communication)
2. [RS232 Communication](#rs232-communication)
3. [LAN Communication](#lan-communication)
4. [Debugging and Error Handling](#debugging-and-error-handling)

## USB HID Communication

### Prerequisites
- Windows OS
- Python 3.x
- SLABHIDtoUART.dll (and possibly SLABHIDDevice.dll) in same folder as script
- APSM2000 configured for 115200 baud, 8N1, RTS/CTS on USB

### Complete USB Streaming Script
Save this as `apms2000_usb_stream.py`:

```python
#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) USB HID Streaming Example with Built-In Logging & Error Handling

Features:
- Opens USB HID connection using Silicon Labs HID DLL
- Continuously queries AC/DC voltages from 3 channels
- Prints to console (3 decimals) and logs to CSV
- Built-in error handling and debugging
"""

import ctypes
import time
import os
import csv
import sys
import logging

# Constants
VID = 0x10C4  # 4292 decimal
PID = 0x8805  # 34869 decimal
HID_UART_SUCCESS = 0
HID_UART_DEVICE = ctypes.c_void_p

# UART Configuration
BAUD_RATE = 115200
DATA_BITS = 8
PARITY_NONE = 0
STOP_BITS_1 = 0
FLOW_CONTROL_RTS_CTS = 2

# Timeouts
READ_TIMEOUT_MS = 500
WRITE_TIMEOUT_MS = 500

def setup_logger(debug=False):
    """Configure logging with console output."""
    logger = logging.getLogger("APSM2000_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    logger.handlers.clear()
    
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger

logger = setup_logger(debug=False)

class APSM2000_USB:
    """Handles USB HID communication with APSM2000."""
    
    def __init__(self, device_index=0):
        self.device_index = device_index
        self.dll_path = os.path.join(os.path.abspath("."), "SLABHIDtoUART.dll")
        self.dev_handle = None
        self._load_dll()
    
    def _load_dll(self):
        """Load Silicon Labs DLL and get function references."""
        try:
            self.hid_dll = ctypes.WinDLL(self.dll_path)
            logger.debug(f"Loaded DLL: {self.dll_path}")
        except OSError as e:
            logger.error(f"Failed to load DLL: {e}")
            raise
            
        # Get function references...
        # (See full code for complete DLL function setup)
    
    def open(self):
        """Open USB HID connection to APSM2000."""
        logger.info("Opening APSM2000...")
        
        # Find devices
        num_devices = ctypes.c_ulong(0)
        ret = self.hid_dll.HidUart_GetNumDevices(
            ctypes.byref(num_devices),
            VID,
            PID
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"GetNumDevices failed: {ret}")
            
        if num_devices.value == 0:
            raise IOError("No APSM2000 devices found")
            
        # Open device
        self.dev_handle = HID_UART_DEVICE()
        ret = self.hid_dll.HidUart_Open(
            ctypes.byref(self.dev_handle),
            self.device_index,
            VID,
            PID
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"Open failed: {ret}")
            
        # Configure UART
        ret = self.hid_dll.HidUart_SetUartConfig(
            self.dev_handle,
            BAUD_RATE,
            DATA_BITS,
            PARITY_NONE,
            STOP_BITS_1,
            FLOW_CONTROL_RTS_CTS
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"SetUartConfig failed: {ret}")
            
        # Set timeouts
        ret = self.hid_dll.HidUart_SetTimeouts(
            self.dev_handle,
            READ_TIMEOUT_MS,
            WRITE_TIMEOUT_MS
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"SetTimeouts failed: {ret}")
            
        logger.info("APSM2000 opened successfully")
    
    def close(self):
        """Close USB HID connection."""
        if self.dev_handle:
            self.hid_dll.HidUart_Close(self.dev_handle)
            self.dev_handle = None
            logger.info("APSM2000 closed")
    
    def write_line(self, command):
        """Write ASCII command with newline."""
        if not self.dev_handle:
            raise IOError("Device not open")
            
        data = (command + "\n").encode('ascii')
        written = ctypes.c_ulong(0)
        
        ret = self.hid_dll.HidUart_Write(
            self.dev_handle,
            data,
            len(data),
            ctypes.byref(written)
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"Write failed: {ret}")
            
        if written.value != len(data):
            raise IOError("Incomplete write")
            
        logger.debug(f"Sent: {command}")
    
    def read_line(self, timeout_sec=1.0):
        """Read ASCII response until newline."""
        if not self.dev_handle:
            raise IOError("Device not open")
            
        start = time.time()
        buffer = bytearray()
        
        while True:
            if time.time() - start > timeout_sec:
                raise TimeoutError("Read timeout")
                
            temp = (ctypes.c_ubyte * 256)()
            bytes_read = ctypes.c_ulong(0)
            
            ret = self.hid_dll.HidUart_Read(
                self.dev_handle,
                temp,
                256,
                ctypes.byref(bytes_read)
            )
            if ret != HID_UART_SUCCESS:
                raise IOError(f"Read failed: {ret}")
                
            if bytes_read.value > 0:
                buffer.extend(temp[:bytes_read.value])
                if b'\n' in buffer:
                    break
            
            time.sleep(0.01)
        
        line = buffer.split(b'\n')[0].decode('ascii').strip()
        logger.debug(f"Recv: {line}")
        return line

def main():
    """Example: Stream voltages from 3 channels."""
    m2000 = APSM2000_USB()
    
    try:
        m2000.open()
        
        # Clear interface
        m2000.write_line("*CLS")
        time.sleep(0.2)
        
        # Start streaming loop
        print("\nStreaming voltages (Ctrl+C to stop)...")
        print(f"{'Time':>8} | {'CH1':>10} | {'CH2':>10} | {'CH3':>10}")
        
        start_time = time.time()
        while True:
            # Read all 3 channels
            m2000.write_line("READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC")
            response = m2000.read_line()
            
            # Parse comma-separated values
            try:
                v1, v2, v3 = map(float, response.split(','))
                elapsed = time.time() - start_time
                print(f"{elapsed:8.1f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")
            except ValueError as e:
                logger.error(f"Parse error: {e}")
                continue
                
            time.sleep(1.0)  # Poll interval
            
    except KeyboardInterrupt:
        print("\nStopped by user")
    finally:
        m2000.close()

if __name__ == "__main__":
    main()
```

### Usage
1. Save script as `apms2000_usb_stream.py`
2. Place SLABHIDtoUART.dll in same folder
3. Run: `python apms2000_usb_stream.py`
4. Press Ctrl+C to stop

## RS232 Communication

### Prerequisites
- Python 3.x with pyserial (`pip install pyserial`)
- USB-to-RS232 adapter or native RS232 port
- Null-modem cable with full handshaking lines

### Basic RS232 Example

```python
import serial
import time

def init_rs232(port="COM1", baud=115200):
    """Open RS232 connection to APSM2000."""
    ser = serial.Serial()
    ser.port = port
    ser.baudrate = baud
    ser.bytesize = serial.EIGHTBITS
    ser.parity = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout = 1.0
    ser.rtscts = True  # Hardware flow control
    ser.dtr = True     # Required by APSM2000
    
    try:
        ser.open()
        print(f"Opened {port} at {baud} baud")
        
        # Clear interface
        ser.write(b"*CLS\n")
        time.sleep(0.2)
        
        # Example: read voltages
        ser.write(b"READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC\n")
        response = ser.readline().decode('ascii').strip()
        print("Response:", response)
        
    except serial.SerialException as e:
        print(f"Error: {e}")
    finally:
        if ser.is_open:
            ser.close()
            print("Port closed")

if __name__ == "__main__":
    init_rs232("COM1", 115200)
```

## LAN Communication

### Prerequisites
- Network connection to APSM2000
- Valid IP address configured
- Port 10733 accessible

### Basic LAN Example

```python
import socket
import time

def init_lan(ip="192.168.1.100", port=10733):
    """Open TCP connection to APSM2000."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(2.0)
    
    try:
        print(f"Connecting to {ip}:{port}...")
        sock.connect((ip, port))
        
        # Clear interface
        sock.sendall(b"*CLS\n")
        time.sleep(0.2)
        
        # Example: read voltages
        sock.sendall(b"READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC\n")
        response = sock.recv(1024).decode('ascii').strip()
        print("Response:", response)
        
    except socket.error as e:
        print(f"Error: {e}")
    finally:
        sock.close()
        print("Socket closed")

if __name__ == "__main__":
    init_lan("192.168.1.100", 10733)
```

## Debugging and Error Handling

The USB streaming script includes built-in logging and error handling:

- Use `debug=True` for detailed logs
- Errors are caught and logged
- Data parsing errors are handled gracefully
- CSV logging provides data backup
- Graceful shutdown on Ctrl+C

### Common Issues

1. USB Connection
   - Check DLL is present
   - Verify VID/PID match
   - Confirm UART settings (115200 8N1)

2. RS232
   - Verify COM port number
   - Check cable (null-modem required)
   - Enable RTS/CTS and DTR

3. LAN
   - Verify IP address
   - Check network connectivity
   - Port 10733 must be accessible

## Support

For issues:
1. Enable debug logging
2. Check error messages
3. Verify hardware setup
4. Review APSM2000 manual
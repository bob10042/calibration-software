# APSM2000 (M2000) USB Communication Guide

This guide contains all the Python code examples and explanations for communicating with an APSM2000 (M2000 Series) power analyzer via USB, including error handling and data logging.

## Table of Contents
1. Basic USB Script
2. Enhanced Version with Error Handling
3. Complete Version with Data Logging

## Requirements
- Windows PC with Python 3.x
- SLABHIDtoUART.dll (and possibly SLABHIDDevice.dll) in script directory
- APSM2000 configured for USB (115200 baud, 8N1, RTS/CTS)
- Python libraries: ctypes (built-in)

## 1. Basic USB Script
This is the minimal version that opens the USB connection and reads voltages:

```python
#!/usr/bin/env python3
"""
Basic APSM2000 USB Communication Example
"""
import ctypes
import time
import os

# Constants
VID = 0x10C4  # 4292 decimal
PID = 0x8805  # 34869 decimal
HID_UART_SUCCESS = 0
HID_UART_DEVICE = ctypes.c_void_p

def init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False):
    """
    Basic example of communicating with the APSM2000 via USB-RS232.
    """
    ser = serial.Serial()
    ser.port = port_name
    ser.baudrate = baud_rate
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0  # seconds read timeout
    ser.rtscts   = True
    ser.dtr      = True

    try:
        ser.open()
        print(f"Opened {port_name} at {baud_rate} baud.")

        # Clear interface
        ser.write(b"*CLS\n")
        time.sleep(0.1)

        # Request ID
        ser.write(b"*IDN?\n")
        response = ser.readline().decode('ascii').strip()
        print("Raw *IDN? response:", response)

        # Parse fields
        fields = response.split(',')
        print("Parsed fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

    except serial.SerialException as e:
        print("Serial error:", e)
    except Exception as e:
        print("General exception:", e)
    finally:
        if ser.is_open:
            ser.close()
            print(f"Closed {port_name}")

if __name__ == "__main__":
    init_and_read_idn_rs232(port_name="COM4", baud_rate=115200)
```

## 2. Enhanced Version with Error Handling
This version adds robust error handling and logging:

```python
#!/usr/bin/env python3
"""
APSM2000 USB Communication with Error Handling
"""
import ctypes
import time
import os
import logging

# Setup logging
logger = logging.getLogger("APSM2000_Logger")
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(ch)

class APSM2000_USB:
    def __init__(self, device_index=0):
        self.device_index = device_index
        self.dev_handle = None
        self.dll_funcs = None

    def open(self):
        """Opens USB connection with error handling"""
        try:
            # Load DLL
            dll_path = os.path.join(os.path.dirname(__file__), "SLABHIDtoUART.dll")
            self.dll = ctypes.WinDLL(dll_path)
            
            # Get device count
            num_devices = ctypes.c_ulong(0)
            ret = self.dll.HidUart_GetNumDevices(ctypes.byref(num_devices), VID, PID)
            if ret != HID_UART_SUCCESS:
                raise IOError(f"GetNumDevices failed: {ret}")

            if num_devices.value == 0:
                raise IOError("No APSM2000 devices found")

            # Open device
            self.dev_handle = HID_UART_DEVICE()
            ret = self.dll.HidUart_Open(ctypes.byref(self.dev_handle), 
                                      self.device_index, VID, PID)
            if ret != HID_UART_SUCCESS:
                raise IOError(f"Open failed: {ret}")

            logger.info("APSM2000 USB opened successfully")
            return True

        except Exception as e:
            logger.error(f"Error opening APSM2000: {e}")
            if self.dev_handle:
                self.close()
            return False

    def close(self):
        """Closes USB connection safely"""
        if self.dev_handle:
            self.dll.HidUart_Close(self.dev_handle)
            self.dev_handle = None
            logger.info("APSM2000 USB closed")

    def write_command(self, cmd):
        """Writes command with error checking"""
        if not self.dev_handle:
            raise IOError("Device not open")
        
        try:
            cmd_bytes = (cmd + "\n").encode('ascii')
            written = ctypes.c_ulong(0)
            ret = self.dll.HidUart_Write(self.dev_handle, cmd_bytes, 
                                       len(cmd_bytes), ctypes.byref(written))
            if ret != HID_UART_SUCCESS:
                raise IOError(f"Write failed: {ret}")
            if written.value != len(cmd_bytes):
                raise IOError("Incomplete write")
            
            logger.debug(f"Sent: {cmd}")
            return True

        except Exception as e:
            logger.error(f"Write error: {e}")
            return False

    def read_response(self, timeout_sec=1.0):
        """Reads response with timeout"""
        if not self.dev_handle:
            raise IOError("Device not open")

        try:
            buffer = bytearray()
            start_time = time.time()
            
            while (time.time() - start_time) < timeout_sec:
                temp = (ctypes.c_ubyte * 64)()
                bytes_read = ctypes.c_ulong(0)
                
                ret = self.dll.HidUart_Read(self.dev_handle, temp, 64, 
                                          ctypes.byref(bytes_read))
                if ret != HID_UART_SUCCESS:
                    raise IOError(f"Read failed: {ret}")
                
                if bytes_read.value > 0:
                    buffer.extend(temp[:bytes_read.value])
                    if b'\n' in buffer:
                        break
                time.sleep(0.01)
            
            if b'\n' not in buffer:
                raise TimeoutError("No response within timeout")
            
            response = buffer.split(b'\n')[0].decode('ascii').strip()
            logger.debug(f"Received: {response}")
            return response

        except Exception as e:
            logger.error(f"Read error: {e}")
            return None

def main():
    m2000 = APSM2000_USB()
    try:
        if m2000.open():
            # Clear interface
            m2000.write_command("*CLS")
            time.sleep(0.1)
            
            # Get ID
            m2000.write_command("*IDN?")
            response = m2000.read_response()
            if response:
                print(f"Device ID: {response}")
            
            # Read voltages example
            m2000.write_command("READ? VOLTS:CH1:ACDC,VOLTS:CH2:ACDC,VOLTS:CH3:ACDC")
            response = m2000.read_response()
            if response:
                voltages = [float(v) for v in response.split(',')]
                print(f"Voltages: {voltages}")
    
    except Exception as e:
        logger.error(f"Error in main: {e}")
    finally:
        m2000.close()

if __name__ == "__main__":
    main()
```

## 3. Complete Version with Data Logging
For the complete version with data logging and comprehensive error handling, see the full script in this repository: `apms2000_usb_stream.py`

This version includes:
- Continuous voltage streaming
- CSV data logging
- Debug logging
- Comprehensive error handling
- Front panel lockout control

## Usage Tips

1. Basic Connection Test:
```python
m2000 = APSM2000_USB()
m2000.open()
m2000.write_command("*IDN?")
print(m2000.read_response())
m2000.close()
```

2. Reading Voltages:
```python
m2000.write_command("READ? VOLTS:CH1:ACDC,VOLTS:CH2:ACDC,VOLTS:CH3:ACDC")
response = m2000.read_response()
voltages = [float(v) for v in response.split(',')]
```

3. Front Panel Control:
```python
# Lock front panel
m2000.write_command("LOCKOUT")

# Later, restore local control
m2000.write_command("LOCAL")
```

## Troubleshooting

1. DLL Not Found:
- Ensure SLABHIDtoUART.dll is in the same directory as your script
- Or add its location to your system PATH

2. Device Not Found:
- Check USB connection
- Verify APSM2000 USB settings (115200 8N1 RTS/CTS)
- Try unplugging/replugging the device

3. Communication Errors:
- Check for timeout settings
- Verify command syntax
- Look for error messages in the logs

## License
This code is provided as-is for educational and development purposes.

## Support
For issues with the APSM2000 hardware itself, contact the manufacturer's support.
For code issues, check the error logs and verify USB settings match.
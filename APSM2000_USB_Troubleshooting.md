# APSM2000 USB Communication Troubleshooting Guide

## Common Issues and Solutions

### 1. Read Timeout Errors
If you're seeing "Read timeout" or "No response received after max attempts":

1. Check Physical Connection
   - Ensure USB cable is firmly connected
   - Try a different USB port
   - Try a different USB cable

2. Verify APSM2000 Settings
   ```
   Check front panel USB settings:
   - Baud Rate: 115200
   - Data Bits: 8
   - Parity: None
   - Stop Bits: 1
   - Flow Control: RTS/CTS
   ```

3. DLL Configuration
   - Verify SLABHIDtoUART.dll is in the correct location
   - Check Windows Device Manager for USB device presence
   - VID (0x10C4) and PID (0x8805) should match

4. Code Adjustments
   ```python
   # Try increasing timeouts
   READ_TIMEOUT_MS = 1000  # Increase to 1 second
   WRITE_TIMEOUT_MS = 1000
   
   # Reduce polling frequency
   poll_interval = 2.0  # Increase to 2 seconds between reads
   
   # Add delay after opening device
   m2000.open()
   time.sleep(2)  # Add 2-second delay after opening
   ```

### 2. Write Errors (Error 17)
If seeing "Error in Write: 17":

1. Close Other Applications
   - Ensure no other programs are using the USB device
   - Only one connection allowed at a time

2. Reset Device
   ```
   1. Close all programs
   2. Unplug USB cable
   3. Wait 10 seconds
   4. Reconnect USB cable
   5. Wait for Windows to enumerate device
   ```

3. Code Fix
   ```python
   def write_command(self, cmd):
       # Add retry logic
       max_retries = 3
       for attempt in range(max_retries):
           try:
               # Flush before writing
               self.dll.HidUart_FlushBuffers(self.dev_handle, True, True)
               time.sleep(0.1)
               
               # Write command
               cmd_bytes = (cmd + "\n").encode('ascii')
               written = ctypes.c_ulong(0)
               ret = self.dll.HidUart_Write(self.dev_handle, 
                                          cmd_bytes, 
                                          len(cmd_bytes), 
                                          ctypes.byref(written))
               if ret == HID_UART_SUCCESS:
                   return True
               
               # If failed, wait before retry
               time.sleep(0.5)
           except Exception as e:
               print(f"Write attempt {attempt + 1} failed: {e}")
       
       return False
   ```

### 3. Device Not Found
If the script can't find the APSM2000:

1. Check Device Manager
   - Look under "Human Interface Devices"
   - Should see "USB Input Device" with VID_10C4&PID_8805

2. Driver Installation
   ```
   1. Download latest Silicon Labs CP210x VCP driver
   2. Uninstall existing driver
   3. Reboot PC
   4. Install new driver
   5. Reconnect device
   ```

3. Code Verification
   ```python
   # Add debug prints
   num_devices = ctypes.c_ulong(0)
   ret = self.dll.HidUart_GetNumDevices(ctypes.byref(num_devices), 
                                       VID, PID)
   print(f"GetNumDevices returned: {ret}")
   print(f"Found {num_devices.value} devices")
   ```

### 4. Data Corruption
If receiving garbled data:

1. Buffer Management
   ```python
   def read_response(self, timeout_sec=1.0):
       # Clear input buffer first
       self.dll.HidUart_FlushBuffers(self.dev_handle, False, True)
       
       # Read with validation
       buffer = bytearray()
       while True:
           chunk = self._read_chunk()
           if not chunk:
               break
           
           # Validate ASCII data
           try:
               chunk.decode('ascii')
               buffer.extend(chunk)
           except UnicodeDecodeError:
               print("Received non-ASCII data, skipping")
               continue
           
           if b'\n' in buffer:
               break
       
       return buffer.decode('ascii').strip()
   ```

2. Flow Control
   - Ensure RTS/CTS is enabled both in code and device
   - Try reducing baud rate temporarily to test

### 5. Performance Optimization
If communication is slow:

1. Batch Commands
   ```python
   # Instead of multiple single reads
   m2000.write_command("READ? VOLTS:CH1:ACDC,VOLTS:CH2:ACDC,VOLTS:CH3:ACDC")
   ```

2. Buffer Sizes
   ```python
   # Adjust buffer sizes
   WRITE_BUFFER_SIZE = 64
   READ_BUFFER_SIZE = 64
   ```

3. Polling Interval
   ```python
   # Balance between responsiveness and CPU usage
   POLL_INTERVAL = 0.5  # 500ms
   ```

## Diagnostic Tools

### 1. Test Basic Communication
```python
def test_communication(m2000):
    """Basic communication test"""
    try:
        # 1. Clear interface
        m2000.write_command("*CLS")
        time.sleep(0.5)
        
        # 2. Try ID query
        m2000.write_command("*IDN?")
        response = m2000.read_response(timeout_sec=2.0)
        if response:
            print("Communication OK:", response)
            return True
    except Exception as e:
        print("Test failed:", e)
    return False
```

### 2. USB Monitor Script
```python
def monitor_usb_status():
    """Monitor USB device presence"""
    while True:
        try:
            num_devices = ctypes.c_ulong(0)
            ret = dll.HidUart_GetNumDevices(ctypes.byref(num_devices), 
                                          VID, PID)
            print(f"Devices found: {num_devices.value}")
            time.sleep(1)
        except KeyboardInterrupt:
            break
```

## Best Practices

1. Always close device properly
   ```python
   try:
       # Your code here
   finally:
       if m2000:
           m2000.close()
   ```

2. Use error handling
   ```python
   try:
       response = m2000.read_response()
   except TimeoutError:
       print("Read timed out")
   except IOError as e:
       print(f"IO Error: {e}")
   except Exception as e:
       print(f"Unexpected error: {e}")
   ```

3. Implement logging
   ```python
   import logging
   logging.basicConfig(level=logging.DEBUG,
                      format='%(asctime)s - %(levelname)s - %(message)s')
   ```

4. Regular validation
   ```python
   def validate_response(response):
       if not response:
           return False
       try:
           # Check for expected format
           values = [float(v) for v in response.split(',')]
           return len(values) == 3  # Expecting 3 channels
       except ValueError:
           return False
   ```

## Additional Resources

1. Silicon Labs Documentation
   - CP210x USB to UART Bridge API
   - USB HID Interface Specification

2. Windows USB Tools
   - USBView
   - Device Manager
   - USB Tree View

3. Python USB Libraries
   - pyusb
   - pyserial
   - hidapi
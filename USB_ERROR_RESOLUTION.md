# APSM2000 USB Error Resolution Guide

## Current Errors Seen
1. Read timeouts (no response after 40 attempts)
2. Write Error 17
3. Failed voltage readings

## Quick Fix Steps

### 1. Stop All Running Scripts
```bash
# If any scripts are running, stop them with Ctrl+C
# Wait at least 10 seconds before trying again
```

### 2. Reset USB Connection
1. Close all programs using the USB-to-RS232 adapter
2. Physically unplug the USB-to-RS232 adapter
3. Wait 10 seconds
4. Plug the adapter back in
5. Wait for Windows to recognize it

### 3. Check APSM2000 Settings
On the APSM2000 front panel:
1. Navigate to USB settings
2. Verify:
   - Baud Rate: 115200
   - Data Bits: 8
   - Parity: None
   - Stop Bits: 1
   - Flow Control: RTS/CTS

### 4. Modify Script Settings
In the Python script, try these changes:

```python
# 1. Increase timeouts
READ_TIMEOUT_MS = 2000   # 2 seconds
WRITE_TIMEOUT_MS = 2000  # 2 seconds

# 2. Add delay after device open
time.sleep(2)  # 2-second delay after opening device

# 3. Reduce polling frequency
poll_interval = 2.0  # 2 seconds between reads

# 4. Add buffer flush before each write
self.dll.HidUart_FlushBuffers(self.dev_handle, True, True)
time.sleep(0.1)  # Short delay after flush
```

### 5. Test Basic Communication
Run the diagnostic script first:
```bash
python apsm2000_usb_test.py
```

This will test basic USB communication before attempting continuous streaming.

### 6. Common Write Error 17 Fixes

1. Driver Reset:
```
1. Open Device Manager
2. Find the USB-to-Serial adapter
3. Right-click -> Uninstall Device
4. Check "Delete driver software"
5. Click OK
6. Unplug adapter
7. Reboot PC
8. Plug in adapter
9. Let Windows install drivers
```

2. Buffer Management:
```python
# Add to your script
def write_with_retry(self, command, max_retries=3):
    for attempt in range(max_retries):
        try:
            # Clear buffers
            self.dll.HidUart_FlushBuffers(self.dev_handle, True, True)
            time.sleep(0.1)
            
            # Write in smaller chunks
            cmd_bytes = (command + "\n").encode('ascii')
            chunk_size = 32  # Smaller chunks
            written = ctypes.c_ulong(0)
            
            for i in range(0, len(cmd_bytes), chunk_size):
                chunk = cmd_bytes[i:i+chunk_size]
                ret = self.dll.HidUart_Write(self.dev_handle,
                                           chunk,
                                           len(chunk),
                                           ctypes.byref(written))
                if ret != HID_UART_SUCCESS:
                    raise IOError(f"Write failed: {ret}")
                time.sleep(0.05)  # Small delay between chunks
            
            return True
            
        except Exception as e:
            print(f"Write attempt {attempt + 1} failed: {e}")
            time.sleep(0.5)  # Wait before retry
            
    return False
```

### 7. USB-to-RS232 Adapter Specific Issues

1. Check Cable Quality:
   - Use a high-quality USB cable
   - Keep cable length under 2 meters
   - Avoid USB hubs if possible

2. Power Management:
   ```
   1. Open Device Manager
   2. Find USB Serial Port
   3. Properties -> Power Management
   4. Uncheck "Allow computer to turn off this device"
   ```

3. Port Settings:
   ```
   1. Device Manager
   2. USB Serial Port -> Properties
   3. Port Settings -> Advanced
   4. Set "Receive (COM) Buffer" to minimum
   5. Set "Transmit (COM) Buffer" to minimum
   ```

### 8. Testing Steps

1. Basic Communication Test:
```python
# Test just opening and sending *CLS
m2000.write_command("*CLS")
time.sleep(1)
```

2. Simple Read Test:
```python
# Test reading device ID
m2000.write_command("*IDN?")
response = m2000.read_response(timeout_sec=2.0)
print("Device ID:", response)
```

3. Single Channel Test:
```python
# Test reading just one channel
m2000.write_command("READ? VOLTS:CH1:ACDC")
response = m2000.read_response(timeout_sec=2.0)
print("CH1 Voltage:", response)
```

### 9. If Problems Persist

1. Try a different USB port (preferably directly on motherboard)
2. Test with a different USB-to-RS232 adapter if available
3. Check Windows Event Viewer for USB-related errors
4. Consider using direct RS232 or LAN connection instead

### 10. Logging for Troubleshooting

Add this to your script for better error tracking:

```python
import logging

logging.basicConfig(
    filename='apsm2000_debug.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Add log statements
logging.debug(f"Write attempt: {command}")
logging.debug(f"Read response: {response}")
logging.error(f"Error occurred: {error}")
```

## Support Resources

1. Check Silicon Labs CP210x driver version
2. Verify Windows USB Serial support
3. Consider APSM2000 firmware version
4. Contact manufacturer support if hardware issue suspected
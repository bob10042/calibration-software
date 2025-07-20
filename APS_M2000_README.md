# APS M2000 Power Analyzer Communication Guide

## Overview
This document provides information about communicating with the APS M2000 Power Analyzer using USB HID interface. The device uses Silicon Labs CP210x USB to UART Bridge for communication.

## Hardware Requirements
- APS M2000 Power Analyzer
- USB Cable
- Computer with USB port

## Software Requirements
- Silicon Labs CP210x USB to UART Bridge VCP Drivers
- Required DLLs (located in USB DLLs and Headers folder):
  * SLABHIDDevice.dll
  * SLABHIDtoUART.dll

## Device Identification
- Vendor ID (VID): 0x10C4 (4292 decimal)
- Product ID (PID): 0xEA80 (60032 decimal)
- Interface: USB HID

## Communication Settings
- Baud Rate: 9600 (default) or 115200
- Data Bits: 8
- Parity: None
- Stop Bits: 1
- Flow Control: None
- Line Endings: \r\n (CRLF)

## Initialization Sequence
1. Load required DLLs
2. Find and open device using VID/PID
3. Enable UART interface
4. Configure UART settings
5. Set timeouts (1000ms recommended)
6. Send initialization commands:
   ```
   *CLS    # Clear device state
   *RST    # Reset device
   REMOTE  # Set remote mode
   ```

## Command Protocol
- Commands must end with CRLF (\r\n)
- Maximum message size: 64 bytes
- First byte of each message must be 0x00 (Report ID)
- Allow sufficient delay between commands (100-500ms)

## Common Commands
```
*IDN?           # Get device identification
*CLS            # Clear device state
*RST            # Reset device
REMOTE          # Enter remote mode
*ERR?           # Query error status
SYST:COMM:STAT? # Query communication status
```

## Implementation Notes
1. Device Connection:
   - Device appears as HID device
   - Uses Silicon Labs CP210x for USB-UART bridge
   - Requires proper driver installation

2. Communication:
   - Use blocking mode for reliable communication
   - Include Report ID (0x00) at start of each message
   - Implement proper delays between commands
   - Handle multiple read attempts for responses

3. Error Handling:
   - Check return values from all operations
   - Implement timeouts for read operations
   - Handle device disconnection gracefully

## Example Implementation
See `simple_usb.py` for a working implementation that demonstrates:
- Device detection and connection
- UART configuration
- Command/response handling
- Error checking
- Resource cleanup

## Troubleshooting
1. No Device Found:
   - Verify device is powered on
   - Check USB connection
   - Confirm driver installation
   - Look for device in Device Manager

2. Communication Errors:
   - Verify baud rate settings
   - Check command format and line endings
   - Ensure proper delays between commands
   - Verify remote mode is set

3. No Response:
   - Check if device is in remote mode
   - Verify command syntax
   - Implement multiple read attempts
   - Check error status with *ERR?

## Reference Files
- `simple_usb.py`: Main implementation example
- `simple_usb_backup.py`: Backup of working implementation
- `APS_M2000_HID_README.md`: Detailed HID protocol documentation
- `APS_M2000_USB_TROUBLESHOOTING.md`: Additional troubleshooting guidance

For further assistance, refer to the APS M2000 manual or contact technical support.

# APS M2000 Power Analyzer HID Communication Protocol

## Overview
This document describes the USB HID (Human Interface Device) communication protocol used to interface with the APS M2000 Power Analyzer. The device uses a USB-HID interface with specific vendor ID (VID) and product ID (PID) for identification.

## Device Identification
- **Vendor ID (VID)**: 0x10C4
- **Product ID (PID)**: 0xEA80
- **Interface**: USB HID
- **Manufacturer**: APS
- **Product**: M2000

## Communication Protocol Details

### Message Format
- **Report ID**: 0x00 (prepended to all messages)
- **Line Endings**: \r\n (CRLF)
- **Maximum Message Size**: 64 bytes (including Report ID)
- **Command Format**: ASCII text commands following SCPI syntax

### Command Structure
```python
message = bytearray([0x00]) + command.encode() + b'\r\n'
```

### Communication Parameters
- **Mode**: Blocking mode recommended for reliable communication
- **Read Timeout**: 1000ms recommended
- **Command Delay**: 0.5s recommended between commands
- **Response Check**: Multiple reads may be necessary (up to 10 attempts)

## Initialization Sequence

1. **Device Detection**
```python
# Auto-detect device using VID/PID
devices = hid.enumerate()
device = find_device(vid=0x10C4, pid=0xEA80)
```

2. **Device Setup**
```python
hid_device = hid.device()
hid_device.open(device['vendor_id'], device['product_id'])
hid_device.set_nonblocking(0)  # Use blocking mode
```

3. **Initial Commands**
```python
# Clear device state
send_command(hid_device, "*CLS")

# Reset device
send_command(hid_device, "*RST")

# Set remote mode
send_command(hid_device, "SYSTEM:REMOTE")

# Get device identification
send_command(hid_device, "*IDN?")
```

## Command Protocol

### Sending Commands
1. Prepend Report ID (0x00)
2. Encode command string to bytes
3. Append CRLF line ending
4. Send using HID write

```python
def send_command(hid_device, command):
    message = bytearray([0x00]) + command.encode() + b'\r\n'
    hid_device.write(message)
```

### Reading Responses
1. Read 64-byte chunks
2. Skip first byte (Report ID)
3. Accumulate data until newline detected
4. Decode response from bytes to string

```python
def read_response(hid_device):
    response = bytearray()
    for _ in range(10):  # Multiple read attempts
        data = hid_device.read(64, timeout_ms=1000)
        if data:
            response.extend(data[1:])  # Skip report ID
            if b'\n' in response:
                break
        time.sleep(0.1)
    return response.decode().strip()
```

## Common SCPI Commands

- **Device Identification**: `*IDN?`
- **Clear State**: `*CLS`
- **Reset**: `*RST`
- **Set Remote Mode**: `SYSTEM:REMOTE`
- **Operation Complete Query**: `*OPC?`

## Error Handling

1. **Timeout Handling**
   - Implement timeouts on read operations
   - Multiple read attempts for reliable data capture

2. **Response Verification**
   - Check for error prefixes in responses
   - Verify response completeness (newline termination)

3. **Device Status**
   - Use `*OPC?` to verify command completion
   - Check for error states with `SYST:ERR?`

## Best Practices

1. **Initialization**
   - Always clear and reset device state at startup
   - Verify device identification before proceeding
   - Set remote mode for reliable control

2. **Command Timing**
   - Allow sufficient delays between commands
   - Use longer delays after reset operations
   - Implement command queuing if needed

3. **Resource Management**
   - Close HID device when done
   - Handle device disconnection gracefully
   - Implement proper error recovery

## Dependencies
- Python `hid` package for USB HID communication
- Time management for command delays and timeouts

## Example Implementation
See `simple_usb.py` for a complete implementation of this protocol.

# APS M2000 USB Setup Readme

This document provides a detailed guide for setting up USB HID communication with the APS M2000 Power Analyzer to send commands and receive responses using the HID interface.

## 1. System Requirements

Before starting, ensure the following:

### Software:
- `SLABHIDtoUART.dll`
- `SLABHIDDevice.dll`

### Hardware:
- APS M2000 Power Analyzer connected via USB
- USB HID interface enabled

### Device Identifiers:
- **Vendor ID (VID)**: `0x10C4` (4292)
- **Product ID (PID)**: `0xEA80` (60032)

## 2. Communication Setup

### UART Configuration
```python
# Configure UART settings
HID_UART_EIGHT_DATA_BITS = 0x03
HID_UART_NO_PARITY = 0x00
HID_UART_SHORT_STOP_BIT = 0x00
HID_UART_RTS_CTS_FLOW_CONTROL = 0x01

# Configure the device
HidUart_SetUartConfig(device,
    115200,                    # Baud rate
    HID_UART_EIGHT_DATA_BITS,  # Data bits
    HID_UART_NO_PARITY,        # Parity
    HID_UART_SHORT_STOP_BIT,   # Stop bits
    HID_UART_RTS_CTS_FLOW_CONTROL  # Flow control
)
```

### Initialization Sequence
1. Open device with VID/PID
2. Configure UART settings
3. Set timeouts (1000ms recommended)
4. Enable UART interface
5. Flush buffers
6. Begin communication

## 3. Command Protocol

### Message Format
- Commands must end with `\r\n` (CRLF)
- Maximum message size: 64 bytes per HID report
- Response termination: `\n`

### Example Commands
```python
# Clear device state
*CLS\r\n

# Reset device
*RST\r\n

# Set remote mode
SYSTEM:REMOTE\r\n

# Get device ID
*IDN?\r\n
```

## 4. Reading Responses

### Status Checking
Before reading data, check UART status:
```python
# Get UART status
tx_fifo, rx_fifo, error_status, line_break = HidUart_GetUartStatus(device)

# Check if data is available
if rx_fifo > 0:
    # Read data
    ...
```

### Reading Data
1. Read in 64-byte chunks (HID report size)
2. Accumulate data until response terminator found
3. Handle timeouts appropriately

```python
def read_with_timeout(device, timeout_ms=1000):
    """Read data with timeout, checking UART status."""
    start_time = time.time()
    response = bytearray()
    
    while (time.time() - start_time) * 1000 < timeout_ms:
        # Check UART status
        status = get_uart_status(device)
        if status and status['rx_fifo'] > 0:
            # Data available, read it
            buffer = read_chunk(64)  # Read 64-byte chunk
            if buffer:
                response.extend(buffer)
                if b'\n' in response:  # Found end of response
                    break
        time.sleep(0.01)  # Prevent busy waiting
    
    return response if response else None
```

## 5. Error Handling

### Common Issues
1. No response to commands:
   - Verify UART is enabled
   - Check flow control settings
   - Ensure proper line endings (\r\n)

2. Incomplete reads:
   - Use status checking before reads
   - Read in chunks (64 bytes)
   - Implement proper timeouts

3. Communication errors:
   - Flush buffers before commands
   - Check error status
   - Verify baud rate settings

### Error Checking
```python
# Query any errors
*ERR?\r\n

# Check operation complete
*OPC?\r\n
```

## 6. Best Practices

1. **Initialization**
   - Always enable UART before communication
   - Clear device state with *CLS
   - Set remote mode for reliable control

2. **Command Timing**
   - Allow 100-200ms between commands
   - Use longer delays after reset operations
   - Check device status before reads

3. **Buffer Management**
   - Flush buffers before critical operations
   - Read all available data from device
   - Handle incomplete reads properly

4. **Error Recovery**
   - Check error status regularly
   - Implement command timeouts
   - Reset device if communication fails

## 7. Python Implementation

See `simple_usb.py` for a complete implementation example that demonstrates:
- Device detection and configuration
- UART setup and management
- Command/response handling
- Error checking and recovery
- Buffer management
- Timeout handling

## 8. Troubleshooting

1. **Device Not Found**
   - Verify VID/PID values
   - Check USB connection
   - Ensure device is powered on

2. **No Response to Commands**
   - Verify UART is enabled
   - Check baud rate configuration
   - Ensure proper line endings
   - Verify flow control settings

3. **Read Errors**
   - Check UART status before reading
   - Implement proper timeouts
   - Read in correct chunk sizes
   - Handle partial reads properly

4. **Write Errors**
   - Verify message format
   - Check command syntax
   - Ensure proper line endings
   - Flush buffers before writing

For further assistance, refer to the APS M2000 manual or contact technical support.

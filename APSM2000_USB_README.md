# APSM2000 (M2000 Series) USB Communications Guide

This repository contains Python scripts for communicating with an APSM2000 (M2000 Series) Power Analyzer over USB HID, RS232 (including USB-to-RS232), or LAN interfaces.

## Requirements

- Windows OS (for USB HID via Silicon Labs DLL)
- Python 3.x
- SLABHIDtoUART.dll (and possibly SLABHIDDevice.dll) in the same folder or system PATH
- APSM2000 configured for:
  - USB: 115200 baud, 8N1, RTS/CTS flow control
  - RS232: Matching baud rate and handshake settings
  - LAN: Valid IP address, port 10733

## Files

- `apms2000_usb_stream.py`: Main streaming script with built-in logging & error handling
- `apms2000_rs232.py`: RS232 communication example (including USB-to-RS232)
- `apms2000_lan.py`: LAN (TCP/IP) communication example

## Quick Start

1. Ensure your APSM2000 is properly configured for USB communications
2. Place the Silicon Labs DLL files in the same folder as the scripts
3. Run the main streaming script:
   ```bash
   python apms2000_usb_stream.py
   ```
4. Data will be logged to apms2000_datalog.csv and displayed in the console
5. Press Ctrl+C to stop

## Features

- Continuous streaming of AC/DC voltages from 3 channels
- Console output formatted to 3 decimal places
- CSV data logging with timestamps
- Built-in error handling and debug logging
- Support for front panel lockout/local control

## Debugging

- Set DEBUG_MODE = True in the script for detailed logging
- Check the console for warnings/errors
- The CSV file is flushed after each write to prevent data loss

## Common Commands

- *CLS: Clear interface
- LOCKOUT: Lock front panel
- LOCAL: Return to local control
- READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC: Read voltages

## Notes

- Only one USB HID connection is allowed at a time
- The APSM2000 automatically enters REMOTE mode upon receiving any valid command
- Make sure handshaking (RTS/CTS) matches between your code and the APSM2000's settings
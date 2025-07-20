# APS M2000 USB Communication Code Status

## Current Version: 24-12-2024

### Implementation Status

1. **Basic Functionality**
   - ✅ USB Device Detection
   - ✅ Device Connection
   - ✅ UART Configuration
   - ✅ Basic Command Sending
   - ✅ Mode Setting (MULTIVPA)
   - ❌ Voltage Reading (In Progress)

2. **Working Features**
   - Successfully loads required DLLs
   - Detects and opens device connection
   - Configures UART settings (115200 baud, 8 data bits, no parity)
   - Sends initialization commands (*CLS, REMOTE)
   - Sets MULTIVPA mode
   - Configures all three channels

3. **Current Issues**
   - Read timeouts when attempting to get voltage measurements
   - Write timeouts occasionally occur after extended operation
   - Device sometimes disconnects during operation

4. **Implementation Details**
   - Uses SLABHIDDevice.dll and SLABHIDtoUART.dll
   - Implements error handling for common USB communication issues
   - Includes timeout handling and retry logic
   - Supports all 3 channels of the device

### Next Steps

1. **Voltage Reading Improvements**
   - Investigate read timeout issues
   - Test different command sequences for voltage measurements
   - Optimize timing between commands

2. **Stability Improvements**
   - Implement better reconnection handling
   - Add more robust error recovery
   - Test longer running sessions

### Usage Notes

1. **Device Setup**
   - Ensure device is in MULTIVPA mode
   - USB connection must be established before running
   - May need to unplug/replug device if communication fails

2. **Known Workarounds**
   - Replug device if write timeouts occur
   - Ensure proper mode is set on device front panel
   - Allow sufficient delay between commands

### Dependencies
- Windows OS
- Python 3.x
- SLABHIDDevice.dll
- SLABHIDtoUART.dll
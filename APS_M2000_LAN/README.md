# APS M2000 LAN Communication

This folder contains scripts and documentation for communicating with the APS M2000 Power Analyzer over LAN.

## Network Configuration

- IP Address: 192.168.15.100
- Port: 10733
- Protocol: TCP/IP
- Line Endings: \r\n (CRLF)

## Command Structure

The communication follows this sequence:

1. **Connection Setup**
   ```
   TCP Connection to 192.168.15.100:10733
   ```

2. **Device Identification**
   ```
   *IDN?
   ```
   Expected Response: Device identification string (e.g., "APSM2000/H500/EN/MU284531215")

3. **Remote Mode**
   ```
   REMOTE
   ```
   Sets device to remote control mode

4. **Channel Configuration**
   For each channel (1-3):
   ```
   CHAN <n>           # Select channel n
   VOLT:RANGE 100     # Set voltage range to 100V
   SAVECONFIG         # Save the configuration
   ```

5. **Data Logging**
   ```
   DATALOG 1          # Enable data logging
   ```

6. **Voltage Measurement**
   For each channel (1-3):
   ```
   CHAN <n>           # Select channel n
   MEAS:VOLT?         # Query voltage measurement
   ```

## Important Notes

1. **Command Termination**
   - All commands must end with \r\n (CRLF)
   - Responses are terminated with \r\n

2. **Timing**
   - 500ms delay between commands
   - 1 second delay between measurement cycles

3. **Channel Selection**
   - Must select channel before sending channel-specific commands
   - Valid channels are 1, 2, and 3

4. **Error Handling**
   - Check responses for error messages
   - Log all communication for debugging

## Usage

### Basic Script
Run the basic script with:
```bash
python apsm2000_lan.py
```

The script will:
1. Connect to the device
2. Initialize and configure channels
3. Start continuous voltage measurements
4. Press Ctrl+C to stop measurements

### Documented Voltage Measurement Script
Run the documented voltage measurement script with:
```bash
python apsm2000_lan_voltage_documented.py
```

This enhanced script provides:
1. Comprehensive voltage measurements across all channels
2. Detailed documentation and error handling
3. 15-second measurement cycles per channel
4. Multiple readings with proper stabilization delays
5. Scientific notation parsing for precise measurements
6. Automatic error checking after each operation

Features:
- Full docstring documentation for all methods
- Robust error handling and connection management
- Clear debug output for troubleshooting
- Configurable measurement parameters
- Support for multiple voltage measurement formats

## Log Files

- apsm2000_lan_debug.log: Contains detailed communication logs
- Logs include timestamps, commands sent, and responses received
- Backup files are automatically created with date and time stamps (e.g., apsm2000_lan_voltage_documented_backup_20241229_224500.py)

## Script Versions

1. **apsm2000_lan.py**: Basic communication and measurement script
2. **apsm2000_lan_voltage_documented.py**: Enhanced voltage measurement script with full documentation
3. **apsm2000_lan_voltage.py**: Original voltage measurement implementation

Each version is maintained with regular backups for version control and safety.

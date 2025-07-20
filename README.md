# AGX Communications Setup Guide

This guide provides comprehensive instructions for setting up communications with the AGX Power Source using three different methods: USB-LAN emulation, direct LAN connection, and USB-to-GPIB interface.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Communication Methods](#communication-methods)
  - [USB-LAN Emulation Mode](#usb-lan-emulation-mode)
  - [Direct LAN Connection](#direct-lan-connection)
  - [USB-to-GPIB Interface](#usb-to-gpib-interface)
  - [RS232 Connection](#rs232-connection)
- [Driver Installation](#driver-installation)
- [SCPI Commands Reference](#scpi-commands-reference)
- [Troubleshooting](#troubleshooting)

## Prerequisites

### Hardware Requirements
- AGX Power Source
- USB Cable (Type A to Type B)
- Ethernet LAN Cable (for direct LAN connection)
- USB-to-GPIB Adapter (for GPIB communication)
- RS232 Cable (for serial communication)
- Computer running Windows 10/11 with administrator rights

### Software Requirements
- Virtual COM Driver (for USB communication)
- Network Driver (for LAN emulation)
- IEEE488.1 standard GPIB driver libraries (for GPIB communication)

## Communication Methods

### USB-LAN Emulation Mode

#### AGX Configuration
1. Enable LAN Emulation Mode:
   - Press **SYST** → Navigate to **INTERFACE SETUP**
   - Select **USB Device Interface**
   - Enable **LAN Emulation**
   - Fixed virtual IP: `192.168.123.1`

#### Computer Setup
1. Connect AGX to computer via USB
2. Configure Virtual Network Adapter:
   - Open Network Adapter Settings
   - Set IPv4 Properties:
     ```
     IP Address: 192.168.123.2
     Subnet Mask: 255.255.255.0
     ```
3. Verify connection:
   ```powershell
   ping 192.168.123.1
   ```
4. Access web interface: `http://192.168.123.1`

### Direct LAN Connection

#### AGX Configuration
1. Configure LAN Settings:
   - Press **SYST** → **INTERFACE SETUP** → **LAN Interface**
   - Set **IP Mode** to **Static**
   - Configure:
     ```
     IP Address: 192.168.1.100
     Subnet Mask: 255.255.255.0
     ```

#### Computer Setup
1. Connect AGX directly to computer via Ethernet
2. Configure Network Adapter:
   ```
   IP Address: 192.168.1.2
   Subnet Mask: 255.255.255.0
   ```
3. Verify connection:
   ```powershell
   ping 192.168.1.100
   ```

### USB-to-GPIB Interface

#### AGX Configuration
1. Enable GPIB Interface:
   - Press **SYST** → **INTERFACE SETUP** → **GPIB Interface**
   - Enable GPIB
   - Set GPIB Address (e.g., 1)

#### Computer Setup
1. Install GPIB adapter drivers
2. Connect USB-to-GPIB adapter
3. Verify connection using driver software (e.g., NI MAX)

### RS232 Connection

#### APS M2000 Configuration
1. Configure RS232 Settings:
   - Baud Rate: 9600
   - Data Bits: 8
   - Stop Bits: 1
   - Parity: None
   - Flow Control: None

#### Computer Setup
1. Connect APS M2000 to computer via RS232 cable
2. Identify COM port in Device Manager
3. Use provided Python scripts for communication:
   - `PyScripts/aps_lan_setup.py` - Main communication script (RS232 version)
   - `PyScripts/aps_lan_setup_backup.py` - Backup of LAN version

## Driver Installation

### USB Drivers
1. Connect AGX via USB
2. Automatic installation:
   - Windows should detect and install drivers automatically
3. Manual installation if needed:
   - Navigate to "drivers/Windows" directory
   - Run `Driver_Installer.exe`

### GPIB Drivers
- Install IEEE488.1 standard GPIB driver libraries
- Recommended: NI-488.2 or equivalent

## SCPI Commands Reference

### USB Configuration
```
SYSTem:COMMunicate:USB:VIRTualport[:ENABle] 1    # Enable virtual COM port
SYSTem:COMMunicate:USB:VIRTualport[:ENABle]?     # Query status
```

### LAN Configuration
```
SYSTem:COMMunicate:USB:LAN[:ENABle] 1            # Enable LAN interface
SYSTem:COMMunicate:LAN:STATus?                   # Query LAN status
SYSTem:COMMunicate:LAN:DHCP[:ENABle] 1          # Enable DHCP
SYSTem:COMMunicate:LAN:DHCP:RENEW               # Renew DHCP lease
```

### GPIB Configuration
```
SYSTem:COMMunicate:GPIB:ENABle 1                # Enable GPIB interface
SYSTem:COMMunicate:GPIB:ADDress?                # Query GPIB address
```

## Troubleshooting

### Cannot Ping AGX
1. Verify physical connections
2. Check IP configuration matches AGX settings
3. Restart AGX and computer network adapter
4. Verify no IP conflicts exist on network

### Web Interface Not Loading
1. Confirm correct IP address
2. Try different web browser
3. Clear browser cache
4. Check firewall settings

### GPIB Communication Issues
1. Verify GPIB address matches AGX settings
2. Check driver installation
3. Test with basic SCPI command: `*IDN?`
4. Verify cable connections

### RS232 Communication Issues
1. Verify COM port settings match device configuration
2. Check cable connections
3. Test with basic command: `*IDN?`
4. Verify no other programs are using the COM port

### General Tips
- Always power cycle device after changing interface settings
- Ensure only one communication method is active at a time
- Use `*IDN?` command to test basic communication
- Check Windows Device Manager for proper driver installation

## Python Control Scripts

### AGX Control Script (agx_control.py)
The main control script provides general-purpose control of the AGX power source with robust error handling and retry logic.

Key Features:
- Serial communication with configurable parameters
- Command retry logic with buffer management
- Response parsing and validation
- Safe shutdown procedures

### APS M2000 Scripts
Located in the PyScripts directory:
- `aps_lan_setup.py` - Main script for APS M2000 communication (RS232 version)
- `aps_lan_setup_backup.py` - Backup of original LAN communication version

### DC Mode Testing (dc_test.py)
Dedicated script for DC voltage testing with safety limits and measurement averaging.

Key Features:
- Safety voltage limit of 120V DC
- 10-reading averaging for accurate measurements
- Proper timing delays:
  - 30s initial stabilization
  - 15s between test points
- Test sequence: 10V, 30V, 60V, 90V, 120V

#### DC Mode Commands
```
# Setup DC Mode
:VOLT:MODE,DC                # Set DC mode
:VOLT:ALC,ON                # Enable automatic level control
:CURR:LIM,10                # Set current limit to 10A
:VOLT:LIM:MIN,0             # Set minimum voltage limit
:VOLT:LIM:MAX,120           # Set maximum voltage limit (safety)

# Voltage Control
:VOLT:DC,<value>            # Set DC voltage level
:MEAS:VOLT:DC?              # Measure DC voltage

# Output Control
:OUTP,ON                    # Enable output
:OUTP,OFF                   # Disable output
```

#### Safe Operation Sequence
1. Initialize with `*RST` and `*CLS`
2. Configure DC mode and safety limits
3. Disable output before voltage changes
4. Set new voltage level
5. Enable output and wait for stabilization
6. Take measurements with averaging
7. Disable output after testing
8. Safe shutdown with `*GTL` and `*RST`

### Backup Files
Both scripts maintain backup copies (`agx_control_backup.py` and `dc_test_backup.py`) for safety and version control.

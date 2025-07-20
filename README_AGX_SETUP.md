# AGX Setup Utility

This utility provides a guided setup process for establishing communication with an AGX power source through USB-LAN emulation, direct LAN connection, or GPIB interface.

## Prerequisites

1. Python 3.7 or higher installed on your system
2. Administrator/root privileges (required for network configuration)
3. Required hardware:
   - AGX Power Source
   - USB Cable (Type A to Type B) for USB-LAN mode
   - Ethernet Cable for direct LAN connection
   - USB-to-GPIB Adapter for GPIB communication
4. Required drivers:
   - NI-VISA (for GPIB communication)
   - Virtual COM Driver (included in drivers folder)
   - Network Driver (included in drivers folder)

## Installation

1. Install the required Python packages:
```bash
pip install -r requirements.txt
```

2. For GPIB communication:
   - Install NI-VISA from National Instruments website
   - Install GPIB adapter drivers
   - Connect the USB-to-GPIB adapter

## Usage

1. Run the utility with administrator privileges:
```bash
# Windows (PowerShell Admin)
python agx_setup_utility.py

# Linux/Unix
sudo python3 agx_setup_utility.py
```

2. Follow the interactive prompts to:
   - Select connection type
   - Configure network settings
   - Set up the AGX device
   - Test communication

## Connection Types

### 1. USB-LAN Emulation Mode
- Uses USB connection with virtual network adapter
- AGX IP: 192.168.123.1
- Computer IP: 192.168.123.2

### 2. Direct LAN Connection
- Uses direct Ethernet connection
- AGX IP: 192.168.1.100
- Computer IP: 192.168.1.2

### 3. GPIB Connection
- Uses USB-to-GPIB adapter
- Requires NI-VISA installation
- Configurable GPIB address

## Features

- Automatic network adapter configuration
- Connection testing with ping
- Web interface verification
- SCPI command testing (GPIB mode)
- Comprehensive error handling
- Step-by-step guidance

## Troubleshooting

1. Network Configuration Fails:
   - Ensure you have administrator privileges
   - Check if the selected network adapter is correct
   - Verify no other devices are using the same IP

2. Cannot Connect to AGX:
   - Verify physical connections
   - Check AGX power and interface settings
   - Ensure correct drivers are installed
   - Check firewall settings

3. GPIB Communication Issues:
   - Verify NI-VISA installation
   - Check GPIB address settings
   - Test adapter with NI MAX
   - Verify GPIB cable connections

## Support Files

- `agx_setup_utility.py`: Main utility script
- `requirements.txt`: Python package dependencies
- `README_AGX_SETUP.md`: This documentation file

## Notes

- Keep the AGX powered on during the entire setup process
- Some network adapters may require a restart after configuration
- GPIB communication requires proper termination and addressing
- Web interface access may be blocked by security software

For additional support, refer to the AGX operation manual or contact technical support.

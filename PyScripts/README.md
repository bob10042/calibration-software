# AGX Python Scripts and Communication Setup

This folder contains Python scripts for AGX power supply control and testing, along with communication setup procedures.

## Scripts Overview

### Communication Test Scripts
- `test_gpib.py` - Basic GPIB connection test
- `test_agx_comms.py` - AGX communication testing
- `configure_agx_network.py` - Network configuration utility

### Voltage Control Scripts
- `voltage_test_sequence.py` - Automated AC/DC voltage testing sequence
- `output_ac_voltage.py` - Interactive AC voltage control
- `output_voltage.py` - General voltage output control
- `agx_control.py` - Main AGX control interface
- `dc_test.py` - DC voltage testing utilities

### Utility Scripts
- `convert_ukas_to_csv.py` - Data conversion utility
- `agx_test_configs.py` - Test configuration management
- `run_agx_tests.py` - Test execution framework

## Communication Setup Procedure

### GPIB Setup
1. **Hardware Requirements**
   - USB-to-GPIB Adapter (e.g., National Instruments GPIB-USB-HS)
   - GPIB Cable
   - Driver software (NI-VISA or Keysight I/O Libraries)

2. **AGX Configuration**
   - Press SYST → INTERFACE SETUP → GPIB Interface
   - Enable GPIB Status
   - Set GPIB Address (default: 1)
   - Restart AGX to apply changes

3. **Computer Setup**
   - Install GPIB adapter driver software
   - Connect USB-to-GPIB adapter
   - Connect GPIB cable between adapter and AGX
   - Verify connection using test scripts

### Dependencies
Required Python packages (from requirements.txt):
- pyvisa>=1.13.0
- pyvisa-py>=0.7.0
- requests>=2.31.0
- netifaces>=0.11.0

### SCPI Commands
Common commands used in scripts:
- `*IDN?` - Get device identification
- `:VOLT:MODE,AC/DC` - Set voltage mode
- `:VOLT:AC/DC,<value>` - Set voltage level
- `:MEAS:VOLT:AC/DC?` - Measure voltage
- `:OUTP,ON/OFF` - Control output
- `:FREQ,<value>` - Set frequency (AC mode)

## Usage Examples

### Basic GPIB Test
```python
python test_gpib.py
```

### Voltage Test Sequence
```python
python voltage_test_sequence.py
```
Tests AC and DC voltages from 10V to 120V with stabilization readings.

### Interactive AC Control
```python
python output_ac_voltage.py
```
Provides interactive control of AC voltage output.

## Troubleshooting

1. **GPIB Connection Issues**
   - Verify GPIB address matches device settings
   - Check cable connections
   - Ensure NI-VISA is properly installed
   - Run `test_gpib.py` for diagnostics

2. **Voltage Control Issues**
   - Verify output is enabled
   - Check voltage limits
   - Ensure proper mode (AC/DC) is selected
   - Allow sufficient stabilization time

3. **Common Error Messages**
   - "No GPIB devices found": Check connections and drivers
   - "Timeout error": Increase timeout value in script
   - "Command error": Verify SCPI command syntax

## Safety Notes

- Always start with voltage at 0V
- Use proper voltage limits
- Enable current limiting
- Implement proper shutdown procedures
- Verify connections before enabling output

## Additional Resources

- AGX Series Operation Manual
- NI-VISA Documentation
- SCPI Programming Guide

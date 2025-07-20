# AGX Control Script Improvements

## Overview
This document details the improvements made to the AGX control script to resolve USB-CDC errors and enhance voltage control functionality. The script has been updated to match the command structure of the working 3150Afx implementation.

## File Structure
- `agx_control.py` - Current working version with all improvements
- `agx_control_backup.py` - Backup of the working version
- `agx_control_original.py` - Original version before improvements
- `README_AGX_CONTROL.md` - This documentation file

## Key Improvements

### 1. Command Structure
- Aligned SCPI commands with 3150Afx format
  - Changed from `OUTPut:STATe OFF` to `OUTP,OFF`
  - Simplified voltage setting commands
  - Standardized command format across all operations
- Added complete initialization sequence with proper settings
  - Voltage/current limits
  - Phase configurations
  - Operating modes

### 2. Enhanced Error Handling
- Improved USB-CDC error recovery
  - Automatic port reset on USB errors
  - Longer stabilization delays
  - Re-initialization after recovery
- Retry mechanism for failed commands
  - Up to 3 retry attempts
  - Progressive delay between retries
  - Detailed error reporting

### 3. Device Control Features
- Robust voltage control
  - Automatic verification of output state
  - Measurement validation
  - Auto-adjustment for out-of-range values
- Safety checks
  - Output state verification
  - Voltage ramping
  - Multiple verification attempts

### 4. Safety Improvements
- Safe shutdown sequences
  - Voltage ramping to zero
  - Output disable verification
  - Device reset on completion
- Error recovery
  - Automatic retry on communication failures
  - Safe state restoration
  - Port reset handling

## Usage

1. Basic operation:
```python
python agx_control.py
```

2. The script will:
   - Initialize the device with proper settings
   - Run through voltage test sequence (10V to 100V)
   - Monitor and verify voltages at each step
   - Safely shut down the device

## Error Handling

The script now handles several types of errors:

1. USB-CDC Errors
   - Automatic port reset
   - Re-initialization sequence
   - Multiple retry attempts

2. Communication Timeouts
   - Automatic retry with delay
   - Buffer clearing
   - Connection recovery

3. Voltage Verification
   - Measurement validation
   - Auto-adjustment for discrepancies
   - Tolerance checking (10%)

## Safety Features

1. Initialization
   - Complete device reset
   - Safe mode setup
   - Proper limit configuration

2. Operation
   - Output verification before voltage changes
   - Measurement validation
   - Automatic adjustment

3. Shutdown
   - Voltage ramping to zero
   - Output disable verification
   - Complete device reset

## Voltage Tests

The project includes two voltage test implementations:

### 1. Basic Test Sequence (agx_control.py)
- Fixed three-phase AC voltage test sequence:
  - Range: 10V to 100V AC
  - Step Size: 10V increments
  - Sequence: 10V, 20V, 30V, 40V, 50V, 60V, 70V, 80V, 90V, 100V
- Features:
  - Simultaneous three-phase voltage setting
  - Output verification
  - 10% tolerance checking
  - Auto-adjustment if outside tolerance
  - Safe ramping between points

### 2. Advanced Test System (agx_voltage_test.py)
- Flexible test configuration via CSV files
- Supports multiple operating modes:
  - AC, DC, AC+DC
  - Single-phase or three-phase
  - Configurable frequency (default 50Hz)
- Advanced measurement features:
  - Multiple samples per measurement
  - Statistical analysis (mean, standard deviation)
  - Maximum deviation calculation
  - Detailed CSV result logging
- Test parameters:
  - Configurable voltage levels
  - Adjustable sample count
  - Custom test point naming
  - GPIB address selection

### Safety Features
Both implementations include:
- Output disabled between voltage changes
- Voltage verification before proceeding
- Automatic shutdown on invalid measurements
- Complete shutdown sequence after tests
- Error handling and recovery

## Results

The improved script successfully:
- Handles USB-CDC errors without manual intervention
- Maintains stable voltage control across all test points
- Provides accurate measurements with 10% tolerance checking
- Ensures safe operation and shutdown between tests

The test sequence demonstrates stable three-phase voltage control from 10V to 100V with precise measurements across all phases, suitable for basic calibration and functionality verification.

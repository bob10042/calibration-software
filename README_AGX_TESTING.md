# AGX Testing Framework

This framework provides structured test configurations and procedures for the AGX power supply testing, based on the original Pps_3150Afx.cs implementation.

## Overview

The testing framework consists of three main components:

1. **Configuration Definitions** (`agx_test_configs.py`)
   - Device initialization settings
   - Mode-specific configurations
   - Test flow patterns
   - Measurement methods

2. **Test Runner** (`run_agx_tests.py`)
   - Implements test procedures
   - Handles instrument communication
   - Manages test flow and measurements
   - Provides results collection

3. **Test Modes**
   - Three Phase AC/DC
   - Split Phase AC/DC
   - Single Phase AC/DC

## Test Configurations

### Device Initialization

The Newton's 4th Power Analyzer requires specific initialization:
- Reset and system configuration
- Phase coupling settings (AC+DC)
- Application mode and bandwidth settings
- Data logging and resolution configuration
- Multilog setup for measurements

### Mode Configurations

#### Three Phase AC
- Voltage Range: 0-300V
- Frequency: 400Hz default
- Phase Settings: 120° (Phase 2), 240° (Phase 3)
- Current Limit: 10A

#### Three Phase DC
- Voltage Range: 0-425V
- Phase Settings: 120° (Phase 2), 240° (Phase 3)
- Current Limit: 10A

#### Split Phase AC
- Voltage Range: 0-600V
- Frequency: 400Hz default
- Phase Settings: 180° (Phase 2)
- Current Limit: 10A

#### Split Phase DC
- Voltage Range: 0-850V
- Phase Settings: 180° (Phase 2)
- Current Limit: 10A

#### Single Phase AC/DC
- Voltage Range: 0-850V
- Current Limit: 10A
- Requires manual linking of three phase outputs

## Test Procedures

### Common Test Flow
1. Initialize instruments
2. Configure test mode
3. Allow initial stabilization (30 seconds)
4. For each test point:
   - Set voltage
   - Wait for stabilization (15 seconds)
   - Take measurements (10 samples averaged)
5. Clean up and reset instruments

### Measurement Methods
- AC Voltage: `:MEAS:VOLT:AC[1-3]?`
- DC Voltage: `:MEAS:VOLT:DC[1-3]?`
- Line-Line Voltage: `:MEAS:VLL1?`

### Important Notes

1. **Stabilization Times**
   - Initial mode change: 30 seconds
   - Between measurements: 15 seconds
   - Sample averaging: 10 samples per measurement

2. **Frequency Response**
   - For frequencies above 800Hz, voltage range should be disabled in newer firmware
   - Frequency response tests use fixed 115V output

3. **Single Phase Testing**
   - Requires manual linking of three phase outputs
   - Confirmation prompt before testing begins

4. **Safety Measures**
   - Output disabled during configuration changes
   - Automatic cleanup on test completion/interruption
   - Instruments reset to local control after testing

## Usage Example

```python
from run_agx_tests import AGXTestRunner

# Initialize test runner
runner = AGXTestRunner()

try:
    # Setup instruments
    runner.setup_instruments()
    
    # Run three phase AC test
    ac_test_points = [10, 25, 50, 75, 100, 115, 135, 150, 200, 240, 270, 300]
    results = runner.run_three_phase_ac_test(ac_test_points)
    
    # Process results
    print("Test Results:", results)
    
finally:
    # Cleanup
    runner.cleanup()
```

## Test Points

### AC Voltage Test Points
- Three Phase: 10V, 25V, 50V, 75V, 100V, 115V, 135V, 150V, 200V, 240V, 270V, 300V
- Split Phase: Same as Three Phase
- Single Phase: Same as Three Phase

### DC Voltage Test Points
- Three Phase: 0V, 25V, 50V, 75V, 100V, 120V, 150V, 200V, 250V, 300V, 350V, 400V, 425V
- Split Phase: Same as Three Phase
- Single Phase: Same as Three Phase

## Error Handling

The framework includes comprehensive error handling:
- Instrument communication errors
- Measurement errors
- Configuration errors
- Cleanup errors

Each operation is wrapped in try-except blocks with appropriate error reporting.

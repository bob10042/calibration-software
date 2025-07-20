"""
AGX Test Configurations extracted from Pps_3150Afx.cs
This module documents the various test configurations and initialization procedures
"""

class AGXConfigurations:
    # Measurement Device (Newton's 4th) Configuration
    NEWTON_4TH_INIT = {
        'commands': [
            '*RST',                      # Reset device
            'TRG',                       # Initiate new measurement, reset range and smoothing
            'KEYBOARD,DISABLE',          # Disable keyboard
            'SYST;POWER',               # Select power meter
            'COUPLI,PHASE1,AC+DC',      # Coupling AC+DC for phase 1
            'COUPLI,PHASE2,AC+DC',      # Coupling AC+DC for phase 2
            'COUPLI,PHASE3,AC+DC',      # Coupling AC+DC for phase 3
            'APPLIC,NORMAL,SPEED3',     # Select application mode
            'BANDWI,WIDE',              # Bandwidth @ 3MHz
            'DATALO,DISABLE',           # Datalog disabled
            'RESOLU,HIGH',              # High data resolution
            'EFFICI,0',                 # Efficiency calculation disabled
            'FAST,ON',                  # Disable screen for faster resolutions
            'FREQFI,OFF',               # Turn frequency filter off
            'SPEED,HIGH',               # High speed mode
            'DISPLAY,VOLTAGE',          # Display voltage mode
            'ZOOM,1,3,5,6,7',          # Set zoom level for display
        ],
        'multilog_setup': [
            'MULTIL,0',                # Reset multilogs
            'MULTIL,1,1,50',           # Phase 1 RMS voltage
            'MULTIL,2,2,50',           # Phase 2 RMS voltage
            'MULTIL,3,3,50',           # Phase 3 RMS voltage
            'MULTIL,4,1,1',            # Phase 1 frequency
            'MULTIL,5,1,79',           # Phase-phase RMS voltage
            'MULTIL,6,1,58',           # Phase 1 DC voltage
            'MULTIL,7,2,58',           # Phase 2 DC voltage
            'MULTIL,8,3,58',           # Phase 3 DC voltage
        ]
    }

    # Three Phase AC Configuration
    THREE_PHASE_AC = {
        'mode': 'AC',
        'phase_config': '3-PHASE',
        'commands': [
            'OUTPUT:AUTO,OFF',
            'INIT,OFF',
            '*LLO',
            'OUTP,OFF',
            'PROG:EXEC,0',
            'VOLT:MODE,AC',
            'FORM,3',
            'VOLT:ALC,ON',
            'CURR:LIM,10',
            'WAVEFORM,1',
            'SENS:PATH,0',
            'FREQ:LIM:MIN,45',
            'FREQ:LIM:MAX,1500',
            'FREQ,400',
            'VOLT:AC:LIM:MIN,0',
            'VOLT:AC:LIM:MAX,300',
            'COUPL,DIRECT',
            'PHAS2,120',
            'PHAS3,240',
            'RANG,1',
            'RAMP,0',
            'OUTP,ON'
        ]
    }

    # Three Phase DC Configuration
    THREE_PHASE_DC = {
        'mode': 'DC',
        'phase_config': '3-PHASE',
        'commands': [
            'OUTP,OFF',
            'INIT,OFF',
            '*LLO',
            'PROG:EXEC,0',
            'VOLT:MODE,DC',
            'FORM,3',
            'VOLT:ALC,ON',
            'CURR:LIM,10',
            'WAVEFORM,1',
            'SENS:PATH,0',
            'VOLT:DC:LIM:MIN,0',
            'VOLT:DC:LIM:MAX,425',
            'COUPL,DIRECT',
            'PHAS2,120',
            'PHAS3,240',
            'RANG,1',
            'RAMP,0',
            'OUTP,ON'
        ]
    }

    # Split Phase AC Configuration
    SPLIT_PHASE_AC = {
        'mode': 'AC',
        'phase_config': '2-PHASE',
        'commands': [
            'OUTP,OFF',
            'INIT,OFF',
            '*LLO',
            'PROG:EXEC,0',
            'VOLT:MODE,AC',
            'FORM,2',
            'FREQ,400',
            'VOLT:ALC,ON',
            'CURR:LIM,10',
            'WAVEFORM,1',
            'SENS:PATH,0',
            'FREQ:LIM:MIN,45',
            'FREQ:LIM:MAX,1500',
            'VOLT:LIM:MIN,0',
            'VOLT:LIM:MAX,600',
            'COUPL,DIRECT',
            'PHAS2,180',
            'RANG,1',
            'RAMP,0',
            'OUTP,ON'
        ]
    }

    # Split Phase DC Configuration
    SPLIT_PHASE_DC = {
        'mode': 'DC',
        'phase_config': '2-PHASE',
        'commands': [
            'OUTP,OFF',
            'INIT,OFF',
            '*LLO',
            'PROG:EXEC,0',
            'VOLT:MODE,DC',
            'FORM,2',
            'VOLT:ALC,ON',
            'CURR:LIM,10',
            'WAVEFORM,1',
            'SENS:PATH,0',
            'VOLT:LIM:MIN,0',
            'VOLT:LIM:MAX,850',
            'COUPL,DIRECT',
            'PHAS2,180',
            'RANG,1',
            'RAMP,0',
            'OUTP,ON'
        ]
    }

    # Single Phase AC Configuration
    SINGLE_PHASE_AC = {
        'mode': 'AC',
        'phase_config': '1-PHASE',
        'commands': [
            'OUTP,OFF',
            'INIT,OFF',
            '*LLO',
            'PROG:EXEC,0',
            'VOLT:MODE,AC',
            'FORM,1',
            'VOLT:ALC,ON',
            'CURR:LIM,10',
            'WAVEFORM,1',
            'SENS:PATH,0',
            'VOLT:LIM:MIN,0',
            'VOLT:LIM:MAX,850',
            'COUPL,DIRECT',
            'RANG,1',
            'RAMP,0',
            'OUTP,ON'
        ]
    }

    # Single Phase DC Configuration
    SINGLE_PHASE_DC = {
        'mode': 'DC',
        'phase_config': '1-PHASE',
        'commands': [
            'OUTP,OFF',
            'INIT,OFF',
            '*LLO',
            'PROG:EXEC,0',
            'VOLT:MODE,DC',
            'FORM,1',
            'VOLT:ALC,ON',
            'CURR:LIM,10',
            'WAVEFORM,1',
            'SENS:PATH,0',
            'VOLT:LIM:MIN,0',
            'VOLT:LIM:MAX,850',
            'COUPL,DIRECT',
            'RANG,1',
            'RAMP,0',
            'OUTP,ON'
        ]
    }

    # Test Flow Patterns
    TEST_FLOWS = {
        'three_phase_ac': {
            'setup': ['newton_4th_init', 'three_phase_ac_config'],
            'stabilization_time': 30000,  # 30 seconds for first measurement
            'measurement_delay': 15000,   # 15 seconds between subsequent measurements
            'measurements': ['voltage_ac1', 'voltage_ac2', 'voltage_ac3']
        },
        'three_phase_dc': {
            'setup': ['newton_4th_init', 'three_phase_dc_config'],
            'stabilization_time': 30000,
            'measurement_delay': 15000,
            'measurements': ['voltage_dc1', 'voltage_dc2', 'voltage_dc3']
        },
        'split_phase_ac': {
            'setup': ['newton_4th_init', 'split_phase_ac_config'],
            'stabilization_time': 30000,
            'measurement_delay': 15000,
            'measurements': ['voltage_line_line']
        },
        'split_phase_dc': {
            'setup': ['newton_4th_init', 'split_phase_dc_config'],
            'stabilization_time': 30000,
            'measurement_delay': 15000,
            'measurements': ['voltage_line_line']
        },
        'single_phase_ac': {
            'setup': ['newton_4th_init', 'single_phase_ac_config'],
            'stabilization_time': 30000,
            'measurement_delay': 15000,
            'measurements': ['voltage_ac1']
        },
        'single_phase_dc': {
            'setup': ['newton_4th_init', 'single_phase_dc_config'],
            'stabilization_time': 30000,
            'measurement_delay': 15000,
            'measurements': ['voltage_dc1']
        }
    }

    # Measurement Methods
    MEASUREMENT_METHODS = {
        'voltage_ac1': ':MEAS:VOLT:AC1?',
        'voltage_ac2': ':MEAS:VOLT:AC2?',
        'voltage_ac3': ':MEAS:VOLT:AC3?',
        'voltage_dc1': ':MEAS:VOLT:DC1?',
        'voltage_dc2': ':MEAS:VOLT:DC2?',
        'voltage_dc3': ':MEAS:VOLT:DC3?',
        'voltage_line_line': ':MEAS:VLL1?'
    }

    # Special Notes
    NOTES = {
        'frequency_response': 'For frequencies above 800Hz, voltage range should be disabled in newer firmware',
        'single_phase_setup': 'Requires manual linking of three phase outputs before testing',
        'measurement_averaging': 'All measurements are averaged over 10 samples',
        'stabilization': 'Initial 30 second stabilization required after mode changes'
    }

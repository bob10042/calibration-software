# Calibration Templates

This directory contains calibration templates for the CalTestGUI application.

## Available Templates

1. **APS_M2003_Calibration_Template.csv** - APS M2003 Power Analyzer calibration template
   - AC/DC Voltage calibration (200mV to 1000V)
   - AC/DC Current calibration (200µA to 20A)
   - Power calibration (10W to 10kW)
   - Frequency calibration (50Hz to 10kHz)
   - Harmonic analysis verification
   - Power factor calibration

2. **Fluke_8508A_Calibration_Template.csv** - Fluke 8508A Reference Multimeter template
   - DC Voltage calibration (200mV to 1000V, 8.5-digit precision)
   - AC Voltage calibration (45Hz-1kHz)
   - DC/AC Current calibration (200µA to 20A)
   - Resistance calibration (2Ω to 20GΩ, 2-wire and 4-wire)
   - Frequency calibration (1Hz to 100kHz)
   - Temperature coefficient and stability tests

3. **Keysight_34471A_Calibration_Template.csv** - Keysight 34471A Truevolt DMM template
   - DC Voltage calibration (100mV to 1000V)
   - AC Voltage calibration (20Hz-300kHz)
   - DC/AC Current calibration (100µA to 3A)
   - Resistance calibration (100Ω to 100MΩ, 2-wire and 4-wire)
   - Capacitance calibration (1nF to 100µF)
   - Frequency/Period calibration
   - Temperature measurement (RTD, Thermocouple)
   - Auto-calibration verification

4. **Vitrek_920A_Calibration_Template.csv** - Vitrek 920A Safety Tester template
   - AC/DC Hipot voltage calibration (up to 12kV)
   - AC/DC Hipot current measurement (up to 100mA)
   - Insulation resistance measurement (up to 2000MΩ)
   - Ground bond resistance measurement (up to 600mΩ)
   - Line leakage and touch current measurement
   - Timer accuracy and ramp rate verification
   - Safety interlocks testing

## Usage

1. Select a meter from the CalTestGUI application
2. Choose the corresponding template file when prompted
3. The application will establish communication with the meter
4. Open the template in Excel for data entry and results

## Converting to Excel Format

To convert CSV templates to Excel format:

1. Open the CSV file in Excel
2. Save As → Excel Workbook (.xlsx)
3. Use the .xlsx file with CalTestGUI for full functionality

## Template Structure

Each template includes:
- Instrument information section
- Multiple calibration test sections
- Tolerance and pass/fail criteria
- Environmental conditions recording
- Calibration standards documentation
- Summary and signature sections

## Customization

Templates can be customized for specific:
- Test procedures
- Tolerance requirements
- Additional test points
- Company-specific formats
- Regulatory compliance needs
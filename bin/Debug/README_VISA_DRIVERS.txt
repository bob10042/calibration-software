CalTestGUI - VISA Drivers Directory
======================================

This directory contains the VISA drivers and runtime libraries needed for 
CalTestGUI to communicate with test instruments.

INCLUDED DRIVERS:
-----------------
✓ Ivi.Visa.dll                    - IVI Foundation VISA.NET interface
✓ NationalInstruments.Visa.dll    - NI VISA.NET wrapper  
✓ NationalInstruments.NI4882.dll  - NI-488.2 GPIB driver
✓ nivisa64.dll                    - NI-VISA runtime (64-bit)
✓ visa32.dll                      - VISA runtime (32-bit compat)
✓ Supporting Visual C++ runtimes

SUPPORTED INTERFACES:
--------------------
✓ USB Test & Measurement Class (USBTMC) instruments
✓ GPIB (IEEE-488) instruments via NI-488.2
✓ Serial/RS-232 instruments via COM ports
✓ Ethernet/TCP-IP instruments

COMMON VISA RESOURCE FORMATS:
-----------------------------
USB:    USB0::0x2A8D::0x0001::MY12345678::INSTR
GPIB:   GPIB0::7::INSTR  
Serial: ASRL1::INSTR (COM1)
TCP-IP: TCPIP0::192.168.1.100::INSTR

FOR FULL SUPPORT:
-----------------
While the included drivers support most common instruments, for complete 
manufacturer support install:

• NI-VISA Runtime (recommended)
  Download: https://www.ni.com/downloads/

• Manufacturer-specific drivers (optional):
  - Keysight IO Libraries Suite
  - Rohde & Schwarz VISA
  - Tektronix OpenChoice

TROUBLESHOOTING:
---------------
If no VISA resources are detected:
1. Install NI-VISA Runtime  
2. Check instrument connections (USB/GPIB)
3. Verify instruments are powered on
4. Run Windows Device Manager to check drivers

For detailed troubleshooting, see VISA_DRIVERS_GUIDE.md in the main directory.

TESTING:
--------
1. Launch CalTestGUI.exe
2. Go to Meters section
3. Click any meter button
4. Check VISA resource dropdown
5. Should show connected instruments

Last Updated: 2025-07-13
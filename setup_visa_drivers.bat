@echo off
echo ================================================================
echo CalTestGUI VISA Drivers Setup
echo ================================================================
echo.

echo Checking current VISA driver installation...
echo.

:: Check if NI-VISA is installed
reg query "HKLM\SOFTWARE\National Instruments\Common\Installer" /v "NIPUBAPI" >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] National Instruments VISA appears to be installed
) else (
    echo [WARNING] NI-VISA not detected in registry
)

:: Check for VISA runtime files
if exist "%SystemRoot%\System32\visa32.dll" (
    echo [OK] VISA runtime files found
) else (
    echo [WARNING] VISA runtime files not found
)

:: Check for included drivers
echo.
echo Checking included CalTestGUI VISA drivers:
if exist "bin\Debug\Ivi.Visa.dll" (
    echo [OK] IVI VISA.NET - bin\Debug\Ivi.Visa.dll
) else (
    echo [MISSING] IVI VISA.NET
)

if exist "bin\Debug\NationalInstruments.Visa.dll" (
    echo [OK] NI VISA.NET - bin\Debug\NationalInstruments.Visa.dll
) else (
    echo [MISSING] NI VISA.NET
)

if exist "bin\Debug\NationalInstruments.NI4882.dll" (
    echo [OK] NI-488.2 GPIB - bin\Debug\NationalInstruments.NI4882.dll
) else (
    echo [MISSING] NI-488.2 GPIB
)

if exist "bin\Debug\nivisa64.dll" (
    echo [OK] NI-VISA Runtime - bin\Debug\nivisa64.dll
) else (
    echo [MISSING] NI-VISA Runtime
)

echo.
echo ================================================================
echo VISA Driver Recommendations:
echo ================================================================
echo.
echo For COMPLETE instrument support, install:
echo.
echo 1. NI-VISA Runtime (2023 Q4 or later)
echo    Download: https://www.ni.com/downloads/
echo    - Provides full USB, Serial, GPIB, Ethernet support
echo    - Required for most professional instruments
echo.
echo 2. Manufacturer-specific drivers (optional):
echo    - Keysight IO Libraries Suite (for Keysight instruments)
echo    - Rohde ^& Schwarz VISA (for R^&S instruments)
echo    - Tektronix OpenChoice (for Tektronix instruments)
echo.
echo For BASIC operation with common USB instruments:
echo    The included drivers in bin\Debug\ should be sufficient
echo.
echo ================================================================
echo Testing VISA Installation
echo ================================================================

:: Try to list VISA resources using included libraries
echo.
echo Testing VISA resource detection...
if exist "CalTestGUI.exe" (
    echo Running CalTestGUI VISA test...
    CalTestGUI.exe /visa-test 2>nul
    if %errorlevel% == 0 (
        echo [OK] VISA resource detection working
    ) else (
        echo [INFO] Launch CalTestGUI to test VISA resource detection
    )
) else (
    echo [INFO] Build CalTestGUI first, then run this script from bin\Debug\
)

echo.
echo ================================================================
echo Setup Complete
echo ================================================================
echo.
echo Next steps:
echo 1. Connect your test instruments
echo 2. Launch CalTestGUI.exe
echo 3. Navigate to Meters section
echo 4. Select a meter type
echo 5. Test VISA resource detection
echo.
echo If no VISA resources are found:
echo - Install NI-VISA Runtime from ni.com
echo - Check instrument USB/GPIB connections
echo - Verify instrument power and settings
echo.
pause
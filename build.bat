@echo off
echo Building CalTestGUI...

:: Set Visual Studio MSBuild path
set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

:: Build configuration
set "CONFIG=Debug"
set "PLATFORM=Any CPU"

:: Clean and Build
echo Cleaning previous build...
"%MSBUILD%" CalTestGUIW.sln /t:Clean /p:Configuration=%CONFIG% /p:Platform="%PLATFORM%"

echo.
echo Building solution...
"%MSBUILD%" CalTestGUIW.sln /t:Rebuild /p:Configuration=%CONFIG% /p:Platform="%PLATFORM%"

:: Check build result
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build completed successfully!
    echo.
    echo Application is available at: bin\Debug\CalTestGUI.exe
    echo.
    choice /C YN /M "Would you like to run the application now"
    if errorlevel 2 goto :end
    if errorlevel 1 start "" "bin\Debug\CalTestGUI.exe"
) else (
    echo.
    echo Build failed with error level %ERRORLEVEL%
    pause
)

:end
echo.
echo Build script completed.
pause

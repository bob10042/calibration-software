# AGX Calibration UI Build and Run Script
$ErrorActionPreference = "Stop"

# Kill existing process if running
Get-Process "AGXCalibrationUI" -ErrorAction SilentlyContinue | Stop-Process -Force

# Build and run sequence
Write-Host "Building AGX Calibration UI..." -ForegroundColor Green

try {
    # Clean and build
    dotnet clean
    if ($LASTEXITCODE -ne 0) { throw "Clean failed" }
    
    dotnet restore
    if ($LASTEXITCODE -ne 0) { throw "Restore failed" }
    
    dotnet build --configuration Debug
    if ($LASTEXITCODE -ne 0) { throw "Build failed" }
    
    # Run the application directly instead of using dotnet run
    Write-Host "Starting application..." -ForegroundColor Green
    Start-Process ".\bin\Debug\net6.0-windows\AGXCalibrationUI.exe"
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

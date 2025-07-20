# CalTestGUI Build Script - PowerShell 7
# Advanced build script utilizing PowerShell 7 features

param(
    [string]$Configuration = "Debug",
    [switch]$Clean,
    [switch]$Rebuild,
    [switch]$RunAfterBuild
)

Write-Host "🔧 CalTestGUI Build Script (PowerShell 7)" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# Project information
$projectInfo = @{
    Name = "CalTestGUI"
    ProjectFile = "CalTestGUI.csproj"
    Configuration = $Configuration
    BuildStartTime = Get-Date
}

Write-Host "`n📋 Build Configuration:" -ForegroundColor Yellow
$projectInfo | ConvertTo-Json | Write-Host

# MSBuild path detection
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    "C:\Program Files\dotnet\dotnet.exe"
)

$msbuild = $msbuildPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $msbuild) {
    Write-Host "❌ MSBuild not found!" -ForegroundColor Red
    exit 1
}

Write-Host "`n🔍 Using MSBuild: $msbuild" -ForegroundColor Green

# Build operations
try {
    if ($Clean -or $Rebuild) {
        Write-Host "`n🧹 Cleaning project..." -ForegroundColor Yellow
        & $msbuild $projectInfo.ProjectFile /t:Clean /p:Configuration=$Configuration
        if ($LASTEXITCODE -ne 0) { throw "Clean failed" }
    }

    Write-Host "`n🔨 Building project..." -ForegroundColor Yellow
    $buildTarget = if ($Rebuild) { "Rebuild" } else { "Build" }
    & $msbuild $projectInfo.ProjectFile /t:$buildTarget /p:Configuration=$Configuration
    
    if ($LASTEXITCODE -eq 0) {
        $buildTime = (Get-Date) - $projectInfo.BuildStartTime
        Write-Host "`n✅ Build successful!" -ForegroundColor Green
        Write-Host "Build time: $($buildTime.TotalSeconds.ToString('F2')) seconds" -ForegroundColor Green
        
        # Check output
        $outputPath = "bin\$Configuration\CalTestGUI.exe"
        if (Test-Path $outputPath) {
            $fileInfo = Get-Item $outputPath
            Write-Host "Output: $outputPath ($(($fileInfo.Length / 1MB).ToString('F2')) MB)" -ForegroundColor Gray
            Write-Host "Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
            
            if ($RunAfterBuild) {
                Write-Host "`n🚀 Running application..." -ForegroundColor Cyan
                Start-Process -FilePath $outputPath
            }
        }
    } else {
        throw "Build failed with exit code $LASTEXITCODE"
    }
} catch {
    Write-Host "`n❌ Build failed: $_" -ForegroundColor Red
    exit 1
}

Write-Host "`n🎉 Build script completed successfully!" -ForegroundColor Green
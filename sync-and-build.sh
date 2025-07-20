#!/bin/bash
echo "=== CalTestGUI Sync and Build Script ==="
echo ""

# Create temp directory if it doesn't exist
echo "Creating Windows temp directory..."
mkdir -p /mnt/c/temp 2>/dev/null

# Sync project to Windows drive
echo "Syncing project files to Windows drive..."
cp -r . /mnt/c/temp/CalTestGUIW/
echo "✓ Files synced to C:\temp\CalTestGUIW"

# Change to Windows directory (for relative paths in build script)
echo ""
echo "Building project..."
cd /mnt/c/temp/CalTestGUIW

# Run Windows build
cmd.exe /c "cd /d C:\temp\CalTestGUIW && build.bat"

# Check if build succeeded
if [ -f "bin/Debug/CalTestGUI.exe" ]; then
    echo ""
    echo "✓ Build successful! Copying binaries back..."
    
    # Copy built files back to WSL
    mkdir -p /home/bob43/CalTestGUIW/bin/Debug 2>/dev/null
    cp -r bin/Debug/* /home/bob43/CalTestGUIW/bin/Debug/
    
    echo "✓ Binaries copied to WSL project directory"
    echo ""
    echo "=== Build Complete ==="
    echo "Run with: ./bin/Debug/CalTestGUI.exe"
else
    echo ""
    echo "✗ Build failed - check output above"
    exit 1
fi
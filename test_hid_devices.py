import hid

def list_all_hid_devices():
    """List all HID devices connected to the system."""
    print("Scanning for HID devices...\n")
    devices = hid.enumerate()
    
    if not devices:
        print("No HID devices found")
        return
        
    for device in devices:
        print(f"VID: 0x{device['vendor_id']:04X} ({device['vendor_id']})")
        print(f"PID: 0x{device['product_id']:04X} ({device['product_id']})")
        print(f"Product: {device.get('product_string', 'Unknown')}")
        print(f"Manufacturer: {device.get('manufacturer_string', 'Unknown')}")
        print(f"Serial Number: {device.get('serial_number', 'Unknown')}")
        print(f"Path: {device['path'].decode() if isinstance(device['path'], bytes) else device['path']}")
        print("-" * 60)

if __name__ == "__main__":
    list_all_hid_devices()

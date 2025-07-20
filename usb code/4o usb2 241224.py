import ctypes
import os
import sys
import time
from ctypes import *

###############################################################################
# 1) DLL Loading
###############################################################################

def get_dll_path():
    """Return x64 or x86 DLL folder based on Python's architecture."""
    if sys.maxsize > 2**32:  # 64-bit
        return os.path.join("USB DLLs and Headers", "x64")
    else:  # 32-bit
        return os.path.join("USB DLLs and Headers", "x86")

def load_dlls():
    """Load SLAB HID/UART DLLs from the appropriate folder."""
    dll_path = get_dll_path()
    if not os.path.exists(dll_path):
        print(f"Error: DLL path does not exist: {dll_path}")
        sys.exit(1)

    try:
        slabhid_dll = ctypes.WinDLL(os.path.join(dll_path, "SLABHIDDevice.dll"))
        slabuart_dll = ctypes.WinDLL(os.path.join(dll_path, "SLABHIDtoUART.dll"))
        print("Successfully loaded SLAB HID UART DLLs")
        return slabhid_dll, slabuart_dll
    except OSError as e:
        print(f"Error loading DLLs: {e}")
        sys.exit(1)

slabhid_dll, slabuart_dll = load_dlls()

###############################################################################
# 2) HID UART Function Prototypes & Constants
###############################################################################

HID_UART_SUCCESS = 0x00
HID_UART_READ_TIMED_OUT = 0x12

# 8N1, no flow control, etc. Adjust if your M2000 is set differently
HID_UART_EIGHT_DATA_BITS = 0x03
HID_UART_NO_PARITY       = 0x00
HID_UART_SHORT_STOP_BIT   = 0x00
HID_UART_NO_FLOW_CONTROL  = 0x00

# Adjust these to match your actual VID/PID if needed
VID = 0x10C4  # 4292 decimal
PID = 0xEA80  # 60032 decimal

# Declare argument/return types for the DLL functions we need
slabuart_dll.HidUart_GetNumDevices.argtypes = [POINTER(c_ulong), c_ushort, c_ushort]
slabuart_dll.HidUart_GetNumDevices.restype  = c_int

slabuart_dll.HidUart_Open.argtypes = [POINTER(c_void_p), c_ulong, c_ushort, c_ushort]
slabuart_dll.HidUart_Open.restype  = c_int

slabuart_dll.HidUart_Close.argtypes = [c_void_p]
slabuart_dll.HidUart_Close.restype  = c_int

slabuart_dll.HidUart_SetUartConfig.argtypes = [
    c_void_p, c_ulong, c_ubyte, c_ubyte, c_ubyte, c_ubyte
]
slabuart_dll.HidUart_SetUartConfig.restype  = c_int

slabuart_dll.HidUart_SetTimeouts.argtypes = [c_void_p, c_ulong, c_ulong]
slabuart_dll.HidUart_SetTimeouts.restype  = c_int

slabuart_dll.HidUart_Read.argtypes  = [c_void_p, POINTER(c_ubyte), c_ulong, POINTER(c_ulong)]
slabuart_dll.HidUart_Read.restype   = c_int

slabuart_dll.HidUart_Write.argtypes = [c_void_p, POINTER(c_ubyte), c_ulong, POINTER(c_ulong)]
slabuart_dll.HidUart_Write.restype  = c_int

###############################################################################
# 3) Helper Functions
###############################################################################

def check_return(result, action):
    """Check and print if a HID_UART call failed."""
    if result != HID_UART_SUCCESS:
        print(f"Error in {action}: code={result}")
        return False
    return True

def find_and_open_device():
    """Locate the M2000 device and open it for communication."""
    device_count = c_ulong(0)
    result = slabuart_dll.HidUart_GetNumDevices(byref(device_count), VID, PID)
    if not check_return(result, "GetNumDevices"):
        return None

    print(f"Found {device_count.value} devices that match VID=0x{VID:X}, PID=0x{PID:X}")
    if device_count.value == 0:
        return None

    device_handle = c_void_p()
    # Try opening device index 0
    result = slabuart_dll.HidUart_Open(byref(device_handle), 0, VID, PID)
    if not check_return(result, "Open"):
        return None

    print("Device opened successfully")

    # Example: set for 9600 baud, 8N1, no flow control. 
    # If your M2000 is set for another rate, match it here.
    result = slabuart_dll.HidUart_SetUartConfig(
        device_handle,
        9600,  # baud rate
        HID_UART_EIGHT_DATA_BITS,
        HID_UART_NO_PARITY,
        HID_UART_SHORT_STOP_BIT,
        HID_UART_NO_FLOW_CONTROL
    )
    if not check_return(result, "SetUartConfig"):
        slabuart_dll.HidUart_Close(device_handle)
        return None

    # Set timeouts (2s read, 2s write)
    result = slabuart_dll.HidUart_SetTimeouts(device_handle, 2000, 2000)
    if not check_return(result, "SetTimeouts"):
        slabuart_dll.HidUart_Close(device_handle)
        return None

    return device_handle

def send_command(device_handle, command_str):
    """
    Send an ASCII command to the M2000, properly filling the buffer
    from index=0 (not offset by 1).
    """
    # Terminate with CR/LF or LF (M2000 typically handles either).
    command_str += "\n"
    cmd_bytes = command_str.encode('ascii')

    # We'll create a report of up to 64 bytes
    report_data = [0]*64
    for i, b in enumerate(cmd_bytes):
        if i < 64:
            report_data[i] = b

    # Convert to c_ubyte array
    c_report = (c_ubyte * 64)(*report_data)
    bytes_written = c_ulong(0)

    # Perform write
    result = slabuart_dll.HidUart_Write(device_handle, c_report, 64, byref(bytes_written))
    if not check_return(result, "Write"):
        return False

    print(f"Sent command: {command_str.strip()}")
    return True

def read_response(device_handle, max_loops=20):
    """
    Read ASCII response from the M2000, up to 'max_loops' of 64 bytes each.
    We break on newline or on no new data after repeated attempts.
    """
    response = bytearray()

    for _ in range(max_loops):
        buffer = (c_ubyte * 64)()
        bytes_read = c_ulong(0)

        # Attempt read
        result = slabuart_dll.HidUart_Read(device_handle, buffer, 64, byref(bytes_read))
        if result == HID_UART_SUCCESS and bytes_read.value > 0:
            # Append everything from index=0 up to bytes_read
            response.extend(buffer[0 : bytes_read.value])

            # If we detect a newline, parse up to that
            if b'\n' in response:
                break
        else:
            # No data or read timed out, short pause
            time.sleep(0.05)

    if not response:
        print("No response received")
        return ""

    # Split at the first newline
    lines = response.split(b'\n', 1)
    first_line = lines[0].decode(errors='replace').strip()
    return first_line

###############################################################################
# 4) Main Logic
###############################################################################

def main():
    print("Looking for M2000 device over USB HID...")

    device = find_and_open_device()
    if not device:
        print("Failed to open M2000 device.")
        return

    try:
        # Clear interface
        send_command(device, "*CLS")
        time.sleep(0.3)

        # Attempt *IDN?
        send_command(device, "*IDN?")
        idn_response = read_response(device)
        if idn_response:
            print(f"[IDN?] => '{idn_response}'")
        else:
            print("[IDN?] => No response or decode error.")

        # Example: read 3-channel voltages
        # The M2000 syntax is: READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3
        time.sleep(0.3)
        send_command(device, "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3")
        volt_line = read_response(device)
        if volt_line:
            print(f"[3-Chan Voltage Read] => '{volt_line}'")

            # Typically something like: +1.23456E+01,+2.34567E+01,+3.45678E+01
            # We can split by comma and parse floats:
            parts = volt_line.split(',')
            if len(parts) == 3:
                try:
                    v1 = float(parts[0])
                    v2 = float(parts[1])
                    v3 = float(parts[2])
                    print(f"Parsed Voltages: CH1={v1:.3f}, CH2={v2:.3f}, CH3={v3:.3f}")
                except ValueError:
                    print("Failed to parse floats from voltage response.")
            else:
                print("Voltage response did not have 3 comma-separated fields.")
        else:
            print("No voltage response.")

    finally:
        # Close device in all cases
        slabuart_dll.HidUart_Close(device)
        print("Device closed.")

if __name__ == "__main__":
    main()

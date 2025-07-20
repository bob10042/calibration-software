import ctypes
import os
import sys
import time
from ctypes import *

# Load the appropriate DLL based on system architecture
def get_dll_path():
    if sys.maxsize > 2**32:  # 64-bit system
        return os.path.join("USB DLLs and Headers", "x64")
    else:  # 32-bit system
        return os.path.join("USB DLLs and Headers", "x86")

# Load the DLLs
def load_dlls():
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

# Constants from SLABCP2110.h
HID_UART_SUCCESS = 0x00
HID_UART_READ_TIMED_OUT = 0x12

HID_UART_EIGHT_DATA_BITS = 0x03
HID_UART_NO_PARITY = 0x00
HID_UART_SHORT_STOP_BIT = 0x00
HID_UART_NO_FLOW_CONTROL = 0x00

VID = 4292  # 0x10C4
PID = 60032 # 0xEA80

# Define function prototypes
slabuart_dll.HidUart_GetNumDevices.argtypes = [POINTER(c_ulong), c_ushort, c_ushort]
slabuart_dll.HidUart_GetNumDevices.restype = c_int

slabuart_dll.HidUart_Open.argtypes = [POINTER(c_void_p), c_ulong, c_ushort, c_ushort]
slabuart_dll.HidUart_Open.restype = c_int

slabuart_dll.HidUart_Close.argtypes = [c_void_p]
slabuart_dll.HidUart_Close.restype = c_int

slabuart_dll.HidUart_SetUartConfig.argtypes = [c_void_p, c_ulong, c_ubyte, c_ubyte, c_ubyte, c_ubyte]
slabuart_dll.HidUart_SetUartConfig.restype = c_int

slabuart_dll.HidUart_SetTimeouts.argtypes = [c_void_p, c_ulong, c_ulong]
slabuart_dll.HidUart_SetTimeouts.restype = c_int

slabuart_dll.HidUart_Read.argtypes = [c_void_p, POINTER(c_ubyte), c_ulong, POINTER(c_ulong)]
slabuart_dll.HidUart_Read.restype = c_int

slabuart_dll.HidUart_Write.argtypes = [c_void_p, POINTER(c_ubyte), c_ulong, POINTER(c_ulong)]
slabuart_dll.HidUart_Write.restype = c_int

def check_return(result, action):
    if result != HID_UART_SUCCESS:
        print(f"Error in {action}: {result}")
        return False
    return True

def find_and_open_device():
    device_count = c_ulong(0)
    result = slabuart_dll.HidUart_GetNumDevices(byref(device_count), VID, PID)
    if not check_return(result, "GetNumDevices"):
        return None

    print(f"Found {device_count.value} devices")
    if device_count.value == 0:
        return None

    device = c_void_p()
    result = slabuart_dll.HidUart_Open(byref(device), 0, VID, PID)
    if not check_return(result, "Open"):
        return None

    print("Device opened successfully")
    time.sleep(1)  # Add delay after opening device

    result = slabuart_dll.HidUart_SetUartConfig(
        device, 9600, HID_UART_EIGHT_DATA_BITS, HID_UART_NO_PARITY,
        HID_UART_SHORT_STOP_BIT, HID_UART_NO_FLOW_CONTROL
    )
    if not check_return(result, "SetUartConfig"):
        slabuart_dll.HidUart_Close(device)
        return None

    result = slabuart_dll.HidUart_SetTimeouts(device, 5000, 5000)  # Increased timeout
    if not check_return(result, "SetTimeouts"):
        slabuart_dll.HidUart_Close(device)
        return None

    return device

def send_command(device, command):
    report = [0] * 64
    command = command + "\r\n"
    command_bytes = [ord(c) for c in command]

    for i, b in enumerate(command_bytes):
        if i + 1 < 64:
            report[i + 1] = b

    command_bytes = (c_ubyte * 64)(*report)
    bytes_written = c_ulong(0)

    result = slabuart_dll.HidUart_Write(device, command_bytes, 64, byref(bytes_written))
    if not check_return(result, "Write"):
        return None

    print(f"Sent command: {command.strip()}")
    return True
def read_response(device, max_attempts=80):  # Doubled max attempts
    """
    Improved response reader with longer timeout and debugging.
    """
    response = bytearray()
    for attempt in range(max_attempts):
        buffer = (c_ubyte * 64)()
        bytes_read = c_ulong(0)

        result = slabuart_dll.HidUart_Read(device, buffer, 64, byref(bytes_read))
        result = slabuart_dll.HidUart_Read(device, buffer, 64, byref(bytes_read))
        if result == HID_UART_SUCCESS and bytes_read.value > 0:
            response.extend(buffer[1:bytes_read.value])
            if b'\n' in response:
                try:
                    response_str = response.split(b'\n')[0].decode().strip()
                    return response_str
                except UnicodeDecodeError as e:
                    print(f"Error decoding response: {e}")
                    return None
        else:
            print(f"Read timeout (attempt {attempt + 1}/{max_attempts})")
            time.sleep(0.1)

    print("No response received after max attempts")
    return None

def stream_voltages(device):
    print("Streaming voltage from VPA1...")
    try:
        while True:
            send_command(device, "MEAS:VOLT?")
            response = read_response(device)
            if response:
                print(f"VPA1 Voltage: {response} V")
            else:
                print("Failed to read voltage from VPA1")
            time.sleep(1)  # Adjust delay as needed
    except KeyboardInterrupt:
        print("Streaming stopped by user.")

def main():
    print("\nLooking for APS M2000...")
    device = find_and_open_device()

    if not device:
        print("Failed to open device")
        return

    try:
        send_command(device, "*CLS")
        time.sleep(1)  # Increased delay
        send_command(device, "*RST")
        time.sleep(2)  # Increased delay after reset

        # Get device identification using direct read
        print("Requesting device identification...")
        send_command(device, "*IDN?")
        time.sleep(1)  # Wait for device to process command
        
        # Direct read attempt for IDN
        buffer = (c_ubyte * 64)()
        bytes_read = c_ulong(0)
        result = slabuart_dll.HidUart_Read(device, buffer, 64, byref(bytes_read))
        
        if result == HID_UART_SUCCESS and bytes_read.value > 0:
            try:
                idn_str = ''.join(chr(b) for b in buffer[1:bytes_read.value]).strip()
                print(f"Device identified as: {idn_str}")
            except Exception as e:
                print(f"Error decoding IDN response: {e}")
        else:
            print(f"Failed to read IDN response: {result}")
        
        time.sleep(1)

        send_command(device, "MODE VPA1")  # Single VPA mode
        time.sleep(0.5)  # Wait for mode change to take effect

        # Only configure channel 1 in VPA1 mode
        send_command(device, "CHAN 1")  # Configure channel 1
        send_command(device, "VOLT:RANGE 100")
        time.sleep(0.5)  # Wait for configuration to settle
        
        # Save configuration twice to ensure it takes effect
        send_command(device, "SAVECONFIG")
        time.sleep(1)  # Wait between saves
        send_command(device, "SAVECONFIG")
        time.sleep(0.5)  # Wait for save to complete

        stream_voltages(device)
    finally:
        slabuart_dll.HidUart_Close(device)
        print("Device closed")

if __name__ == "__main__":
    main()
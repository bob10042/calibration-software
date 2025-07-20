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
HID_UART_DEVICE_NOT_FOUND = 0x01
HID_UART_INVALID_HANDLE = 0x02
HID_UART_INVALID_DEVICE_OBJECT = 0x03
HID_UART_INVALID_PARAMETER = 0x04
HID_UART_INVALID_REQUEST_LENGTH = 0x05
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

slabuart_dll.HidUart_SetUartEnable.argtypes = [c_void_p, c_ubyte]
slabuart_dll.HidUart_SetUartEnable.restype = c_int

def check_return(result, action):
    error_codes = {
        HID_UART_DEVICE_NOT_FOUND: "Device not found",
        HID_UART_INVALID_HANDLE: "Invalid handle",
        HID_UART_INVALID_DEVICE_OBJECT: "Invalid device object",
        HID_UART_INVALID_PARAMETER: "Invalid parameter",
        HID_UART_INVALID_REQUEST_LENGTH: "Invalid request length",
        HID_UART_READ_TIMED_OUT: "Read timed out",
        0x11: "Write timed out",
        0x12: "Read timed out",
        0x17: "Device disconnected"
    }
    if result != HID_UART_SUCCESS:
        error_msg = error_codes.get(result, f"Unknown error code: {result}")
        print(f"Error in {action}: {error_msg}")
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

    # Set non-blocking mode
    result = slabuart_dll.HidUart_SetUartEnable(device, 1)
    if not check_return(result, "SetUartEnable"):
        slabuart_dll.HidUart_Close(device)
        return None

    # Configure UART
    result = slabuart_dll.HidUart_SetUartConfig(
        device, 115200, HID_UART_EIGHT_DATA_BITS, HID_UART_NO_PARITY,
        HID_UART_SHORT_STOP_BIT, HID_UART_NO_FLOW_CONTROL
    )
    if not check_return(result, "SetUartConfig"):
        slabuart_dll.HidUart_Close(device)
        return None

    print("UART configured successfully")

    # Set timeouts
    result = slabuart_dll.HidUart_SetTimeouts(device, 1000, 1000)
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

def read_response(device, max_attempts=20):
    response = bytearray()
    buffer_size = 256  # Increased buffer size

    print("Waiting for response...")
    for attempt in range(max_attempts):
        buffer = (c_ubyte * buffer_size)()
        bytes_read = c_ulong(0)

        result = slabuart_dll.HidUart_Read(device, buffer, buffer_size, byref(bytes_read))
        
        if result == HID_UART_SUCCESS:
            if bytes_read.value > 0:
                response.extend(buffer[1:bytes_read.value])
                if b'\n' in response or b'\r' in response:
                    try:
                        # Try to decode up to the first newline or carriage return
                        response_str = response.split(b'\n')[0].split(b'\r')[0].decode().strip()
                        print(f"Raw response: {response_str}")
                        return response_str
                    except UnicodeDecodeError:
                        print(f"Error decoding response: {bytes(response)}")
                        return None
            time.sleep(0.2)  # Longer delay between read attempts
        elif result == HID_UART_READ_TIMED_OUT:
            print(f"Read timeout (attempt {attempt + 1}/{max_attempts})")
            time.sleep(0.5)  # Even longer delay after timeout
        else:
            print(f"Read error: {result}")
            return None

    print("No complete response received after maximum attempts")
    return None

def stream_voltage(device):
    try:
        print("\nStarting voltage streaming. Press Ctrl+C to stop.")
        
        # Initial delay after mode setting
        time.sleep(2.0)
        
        # Configure all channels once
        print("\nConfiguring channels...")
        for channel in range(1, 4):
            send_command(device, f"CHAN {channel}")
            time.sleep(1.0)
            send_command(device, "VOLT:RANGE 100")
            time.sleep(1.0)
        
        print("\nStarting measurements...")
        time.sleep(1.0)

        # Measure and log voltages
        while True:
            for channel in range(1, 4):
                if send_command(device, f"CHAN {channel}"):
                    time.sleep(1.0)
                    if send_command(device, "MEAS:VOLT?"):
                        time.sleep(1.0)
                        voltage = read_response(device)
                        if voltage:
                            print(f"Channel {channel} Voltage: {voltage} V")
                        else:
                            print(f"Channel {channel}: No reading")
                    else:
                        print(f"Channel {channel}: Failed to send measurement command")
                else:
                    print(f"Channel {channel}: Failed to select channel")
            
            print("-" * 40)
            time.sleep(2.0)  # Longer delay between cycles
            
    except KeyboardInterrupt:
        print("\nVoltage streaming stopped by user.")

def main():
    print("\nLooking for APS M2000...")
    device = find_and_open_device()

    if not device:
        print("Failed to open device")
        return

    try:
        # Basic initialization sequence
        print("Clearing device...")
        send_command(device, "*CLS")
        time.sleep(0.5)
        
        print("Setting to remote mode...")
        send_command(device, "REMOTE")
        time.sleep(0.5)

        print("Setting to MULTIVPA mode...")
        send_command(device, "MODE MULTIVPA")
        time.sleep(1.0)
        
        # Configure channel 1
        print("Configuring channel 1...")
        send_command(device, "CHAN 1")
        time.sleep(0.2)
        send_command(device, "VOLT:RANGE 100")
        time.sleep(0.2)
        send_command(device, "SAVECONFIG")
        time.sleep(0.2)

        # Start streaming voltages
        stream_voltage(device)
        
    finally:
        slabuart_dll.HidUart_Close(device)
        print("Device closed")

if __name__ == "__main__":
    main()

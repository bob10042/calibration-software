import ctypes
import time

# DLL Path (Replace with actual path to your DLL)
HID_DLL_PATH = "./SLABHIDtoUART.dll"

# Configuration Parameters
VID = 0x10C4  # Vendor ID
PID = 0xEA60  # Product ID
BAUD_RATE = 9600  # Default baud rate
TIMEOUT = 1000  # Timeout in milliseconds

# Load the HID DLL
hid_dll = ctypes.CDLL(HID_DLL_PATH)

# Function Prototypes
hid_dll.HidUart_OpenByIndex.argtypes = [ctypes.POINTER(ctypes.c_void_p), ctypes.c_uint]
hid_dll.HidUart_OpenByIndex.restype = ctypes.c_int

hid_dll.HidUart_SetUartConfig.argtypes = [
    ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint
]
hid_dll.HidUart_SetUartConfig.restype = ctypes.c_int

hid_dll.HidUart_Write.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ubyte), ctypes.c_uint, ctypes.POINTER(ctypes.c_uint)]
hid_dll.HidUart_Write.restype = ctypes.c_int

hid_dll.HidUart_Read.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ubyte), ctypes.c_uint, ctypes.POINTER(ctypes.c_uint)]
hid_dll.HidUart_Read.restype = ctypes.c_int

# Initialize HID Connection
def initialize_hid_connection():
    device = ctypes.c_void_p()
    result = hid_dll.HidUart_OpenByIndex(ctypes.byref(device), 0)
    if result != 0:
        print("Failed to open device.")
        return None
    print("Device opened successfully.")

    # Set UART Configuration
    result = hid_dll.HidUart_SetUartConfig(
        device, BAUD_RATE, 8, 0, 1, 0
    )  # 8 data bits, no parity, 1 stop bit, no flow control
    if result != 0:
        print("Failed to configure UART.")
        return None
    print("UART configured successfully.")

    return device

# Send SCPI Command
def send_scpi_command(device, command):
    if device:
        send_data = (ctypes.c_ubyte * len(command))(*[ord(c) for c in command])
        bytes_written = ctypes.c_uint()
        result = hid_dll.HidUart_Write(device, send_data, len(send_data), ctypes.byref(bytes_written))
        if result == 0:
            print(f"Sent: {command}")
        else:
            print("Failed to send command.")

# Read SCPI Response
def read_scpi_response(device):
    if device:
        rx_buffer = (ctypes.c_ubyte * 1024)()
        bytes_read = ctypes.c_uint()
        result = hid_dll.HidUart_Read(device, rx_buffer, 1024, ctypes.byref(bytes_read))
        if result == 0 and bytes_read.value > 0:
            response = ''.join(chr(rx_buffer[i]) for i in range(bytes_read.value))
            print(f"Received: {response.strip()}")
            return response.strip()
        else:
            print("Failed to read response or no data available.")

# Main Function
def main():
    device = initialize_hid_connection()

    if device:
        # Example SCPI Commands
        commands = ["*IDN?", "MEAS:VOLT:DC?", "MEAS:CURR:DC?"]

        for cmd in commands:
            send_scpi_command(device, cmd + "\n")
            time.sleep(0.5)  # Wait for the device to process
            read_scpi_response(device)

        # Close the connection
        hid_dll.HidUart_Close(device)
        print("Connection closed.")

if __name__ == "__main__":
    main()

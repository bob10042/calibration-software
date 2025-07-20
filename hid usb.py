import ctypes
import time
import os

# These constants are from the M2000 manual / Silicon Labs documentation:
VID = 0x10C4  # 4292 in decimal
PID = 0x8805  # 34869 in decimal

# Load the DLLs (adjust the path as needed)
# Make sure these .dll files are accessible to your Python script
dll_path = os.path.abspath(".")  # Example: same folder
hid_dll = ctypes.WinDLL(os.path.join(dll_path, "SLABHIDtoUART.dll"))
# The above call might differ if you’re using a 64-bit Python vs. 32-bit DLL.

# We need a handle type (pointer)
HID_UART_DEVICE = ctypes.c_void_p

# Function prototypes (simplified)
HidUart_GetNumDevices      = hid_dll.HidUart_GetNumDevices
HidUart_Open               = hid_dll.HidUart_Open
HidUart_Close              = hid_dll.HidUart_Close
HidUart_Read               = hid_dll.HidUart_Read
HidUart_Write              = hid_dll.HidUart_Write
HidUart_SetUartConfig      = hid_dll.HidUart_SetUartConfig
HidUart_SetTimeouts        = hid_dll.HidUart_SetTimeouts
HidUart_FlushBuffers       = hid_dll.HidUart_FlushBuffers

# Return codes:
HID_UART_SUCCESS = 0

def open_usb_device(device_index=0):
    """
    Opens the M2000’s USB HID device using the Silicon Labs HID-to-UART driver.
    Returns a valid device handle if successful.
    """
    num_devices = ctypes.c_ulong(0)
    ret = HidUart_GetNumDevices(ctypes.byref(num_devices), VID, PID)
    if ret != HID_UART_SUCCESS:
        raise IOError("HidUart_GetNumDevices failed")

    if num_devices.value == 0:
        raise IOError("No M2000 devices found via USB HID")

    device_handle = HID_UART_DEVICE()
    ret = HidUart_Open(ctypes.byref(device_handle), device_index, VID, PID)
    if ret != HID_UART_SUCCESS:
        raise IOError(f"HidUart_Open failed with code {ret}")

    # Configure the UART side inside the M2000: 115200, 8N1, RTS/CTS, etc.
    # (115200, 8 bits, no parity, 1 stop bit, hardware flow control)
    ret = HidUart_SetUartConfig(
        device_handle,
        ctypes.c_ulong(115200),  # baud rate
        ctypes.c_ubyte(8),       # data bits
        ctypes.c_ubyte(0),       # parity=0 => none
        ctypes.c_ubyte(0),       # stop bits=0 => 1 stop bit
        ctypes.c_ubyte(2)        # flow control=2 => RTS/CTS
    )
    if ret != HID_UART_SUCCESS:
        HidUart_Close(device_handle)
        raise IOError(f"HidUart_SetUartConfig failed with code {ret}")

    # Set timeouts (read=500ms, write=500ms for example)
    ret = HidUart_SetTimeouts(device_handle, 500, 500)
    if ret != HID_UART_SUCCESS:
        HidUart_Close(device_handle)
        raise IOError("HidUart_SetTimeouts failed")

    # Flush any stale buffers
    HidUart_FlushBuffers(device_handle, True, True)

    return device_handle

def close_usb_device(device_handle):
    """Closes the handle to the M2000 USB HID device."""
    if device_handle:
        HidUart_Close(device_handle)

def usb_write(device_handle, cmd_string):
    """
    Writes a string command to the M2000. Must append CR/LF or LF.
    """
    # M2000 typically only needs LF or CR/LF.
    data = (cmd_string + "\n").encode('ascii')
    length = len(data)
    written = ctypes.c_ulong(0)
    ret = HidUart_Write(device_handle, data, length, ctypes.byref(written))
    if ret != HID_UART_SUCCESS:
        raise IOError("HidUart_Write failed")
    if written.value != length:
        raise IOError("Not all bytes were written")

def usb_read_line(device_handle, timeout_sec=2.0):
    """
    Reads a single line from M2000 (until LF) using repeated HidUart_Read calls.
    """
    start = time.time()
    line_bytes = bytearray()

    bufsize = 256
    tempbuf = (ctypes.c_ubyte * bufsize)()
    num_read = ctypes.c_ulong(0)

    while True:
        # Periodically check for timeout
        if (time.time() - start) > timeout_sec:
            raise TimeoutError("USB read timed out waiting for newline")

        # Attempt to read
        ret = HidUart_Read(device_handle, tempbuf, bufsize, ctypes.byref(num_read))
        if ret != HID_UART_SUCCESS:
            raise IOError("HidUart_Read failed")

        if num_read.value > 0:
            # Append to line buffer
            for i in range(num_read.value):
                line_bytes.append(tempbuf[i])
            # Check if we got a LF
            if b'\n' in line_bytes:
                break

        # If no data, small sleep
        time.sleep(0.01)

    # Split at first LF
    line = line_bytes.split(b'\n', 1)[0]
    return line.decode('ascii').strip()

def read_voltages_via_usb():
    """Example usage to read channel voltages over USB HID."""
    device_handle = None
    try:
        device_handle = open_usb_device(device_index=0)

        # Clear interface
        usb_write(device_handle, "*CLS")
        time.sleep(0.2)

        # Request voltages from CH1..CH4
        usb_write(device_handle, "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3, VOLTS:CH4")

        # Read the response line
        response_line = usb_read_line(device_handle, timeout_sec=2.0)
        print("Raw response:", response_line)

        if response_line:
            parts = response_line.split(',')
            voltages = [float(v) for v in parts]
            print("Parsed voltages:", voltages)
        else:
            print("Empty response.")

    except Exception as ex:
        print("Error:", ex)
    finally:
        if device_handle:
            close_usb_device(device_handle)

if __name__ == "__main__":
    read_voltages_via_usb()

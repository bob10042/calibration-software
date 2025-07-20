import ctypes
import os
import sys
import time
from ctypes import *

# Adjust as needed
VID = 0x10C4
PID = 0xEA80
HID_UART_SUCCESS = 0x00

HID_UART_EIGHT_DATA_BITS = 0x03
HID_UART_NO_PARITY       = 0x00
HID_UART_SHORT_STOP_BIT   = 0x00
HID_UART_NO_FLOW_CONTROL  = 0x00

def check_return(result, action):
    if result != HID_UART_SUCCESS:
        print(f"Error in {action}: code={result}")
        return False
    return True

def send_command(device_handle, command_str):
    """Send a line-terminated command (with '\n') at index=0."""
    cmd = command_str + "\n"
    data_bytes = cmd.encode('ascii')
    report_data = [0]*64
    for i,b in enumerate(data_bytes):
        if i<64:
            report_data[i] = b

    c_report = (c_ubyte * 64)(*report_data)
    bytes_written = c_ulong(0)

    result = slabuart_dll.HidUart_Write(device_handle, c_report, 64, byref(bytes_written))
    if not check_return(result, f"Write '{command_str}'"):
        return False
    return True

def read_response(device_handle, max_loops=20):
    """Attempt to read up to 64 bytes multiple times, return line if found."""
    response = bytearray()
    for _ in range(max_loops):
        buf = (c_ubyte * 64)()
        br = c_ulong(0)
        res = slabuart_dll.HidUart_Read(device_handle, buf, 64, byref(br))
        if res == HID_UART_SUCCESS and br.value>0:
            response.extend(buf[0:br.value])
            if b'\n' in response:
                break
        time.sleep(0.05)

    if not response:
        return None
    line = response.split(b'\n',1)[0].decode(errors='replace').strip()
    return line

# load DLL
def load_slab_dlls():
    # adapt path if needed
    import ctypes
    folder = os.path.abspath(".")
    hid_dll = ctypes.WinDLL(os.path.join(folder, "SLABHIDtoUART.dll"))
    return hid_dll

slabuart_dll = load_slab_dlls()

# Declare function prototypes
slabuart_dll.HidUart_GetNumDevices.argtypes = [POINTER(c_ulong), c_ushort, c_ushort]
slabuart_dll.HidUart_GetNumDevices.restype  = c_int

slabuart_dll.HidUart_Open.argtypes = [POINTER(c_void_p), c_ulong, c_ushort, c_ushort]
slabuart_dll.HidUart_Open.restype  = c_int

slabuart_dll.HidUart_Close.argtypes = [c_void_p]
slabuart_dll.HidUart_Close.restype  = c_int

slabuart_dll.HidUart_SetUartConfig.argtypes = [c_void_p, c_ulong, c_ubyte, c_ubyte, c_ubyte, c_ubyte]
slabuart_dll.HidUart_SetUartConfig.restype  = c_int

slabuart_dll.HidUart_SetTimeouts.argtypes = [c_void_p, c_ulong, c_ulong]
slabuart_dll.HidUart_SetTimeouts.restype  = c_int

slabuart_dll.HidUart_Read.argtypes = [c_void_p, POINTER(c_ubyte), c_ulong, POINTER(c_ulong)]
slabuart_dll.HidUart_Read.restype  = c_int

slabuart_dll.HidUart_Write.argtypes= [c_void_p, POINTER(c_ubyte), c_ulong, POINTER(c_ulong)]
slabuart_dll.HidUart_Write.restype = c_int


def main():
    # 1) Find & open
    dev_count = c_ulong(0)
    r = slabuart_dll.HidUart_GetNumDevices(byref(dev_count), VID, PID)
    if not check_return(r, "GetNumDevices"):
        return
    print(f"Found {dev_count.value} device(s).")

    if dev_count.value==0:
        print("No device found.")
        return

    dev = c_void_p()
    r = slabuart_dll.HidUart_Open(byref(dev), 0, VID, PID)
    if not check_return(r, "Open"):
        return
    print("Device opened.")

    # 2) Configure bridging for 115200 8N1 no flow
    # or match your M2000's actual bridging config
    r = slabuart_dll.HidUart_SetUartConfig(
        dev,
        115200,  # change if M2000 bridging is 9600, etc.
        HID_UART_EIGHT_DATA_BITS,
        HID_UART_NO_PARITY,
        HID_UART_SHORT_STOP_BIT,
        HID_UART_NO_FLOW_CONTROL
    )
    if not check_return(r, "SetUartConfig"):
        slabuart_dll.HidUart_Close(dev)
        return

    # 3) Timeouts
    r = slabuart_dll.HidUart_SetTimeouts(dev, 2000, 2000)  # 2s read/write
    if not check_return(r, "SetTimeouts"):
        slabuart_dll.HidUart_Close(dev)
        return

    # 4) Send a simple known command
    time.sleep(0.3)
    send_command(dev, "READ? VOLTS:CH1")
    time.sleep(0.3)
    resp = read_response(dev, max_loops=20)
    if resp is not None:
        print(f"Got response for CH1 volts: '{resp}'")
    else:
        print("No response for CH1 volts")

    # 5) Try 3 channels at once
    time.sleep(0.3)
    send_command(dev, "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3")
    time.sleep(0.3)
    resp2 = read_response(dev, max_loops=20)
    if resp2 is not None:
        print(f"Got 3-channel response: '{resp2}'")
    else:
        print("No response for 3 channels")

    # optional: try IDN last
    time.sleep(0.3)
    send_command(dev, "*IDN?")
    time.sleep(0.3)
    resp3 = read_response(dev, max_loops=20)
    if resp3:
        print(f"IDN => {resp3}")
    else:
        print("No IDN response on USB")

    # 6) Close
    slabuart_dll.HidUart_Close(dev)
    print("Closed device.")

if __name__=="__main__":
    main()

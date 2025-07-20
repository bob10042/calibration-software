"""
APSM2000 (M2000 Series) Python Code Snippets
Contains the three main communication examples (RS232, LAN, USB) as separate functions.
Copy the section you need based on your connection type.
"""

#############################
# RS232 Example (pyserial)
#############################

import time
import serial  # pip install pyserial

def init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False):
    ser = serial.Serial()
    ser.port = port_name
    ser.baudrate = baud_rate
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0
    ser.rtscts = True
    ser.dtr = True

    try:
        ser.open()
        print(f"Opened {port_name} at {baud_rate} baud.")

        ser.write(b"*CLS\n")
        time.sleep(0.1)

        if lock_front_panel:
            ser.write(b"LOCKOUT\n")
            time.sleep(0.1)

        ser.write(b"*IDN?\n")
        response = ser.readline().decode('ascii', errors='replace').strip()
        print("Raw *IDN? response:", response)

        fields = response.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

    except serial.SerialException as e:
        print("Serial error:", e)
    except Exception as e:
        print("General exception:", e)
    finally:
        if ser.is_open:
            ser.close()
            print(f"Closed {port_name}")

#############################
# LAN Example (TCP/IP)
#############################

import socket

def init_and_read_idn_lan(ip_address="192.168.1.100", port=10733, lock_front_panel=False):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(2.0)
    try:
        print(f"Connecting to {ip_address}:{port} ...")
        s.connect((ip_address, port))
        print("Connected.")

        s.sendall(b"*CLS\n")
        time.sleep(0.1)

        if lock_front_panel:
            s.sendall(b"LOCKOUT\n")
            time.sleep(0.1)

        s.sendall(b"*IDN?\n")
        
        response_buffer = []
        while True:
            chunk = s.recv(4096)
            if not chunk:
                break
            response_buffer.append(chunk)
            if b"\n" in chunk:
                break

        raw_response = b"".join(response_buffer).decode('ascii').strip()
        print("Raw *IDN? response:", raw_response)

        fields = raw_response.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

    except socket.timeout:
        print("Socket timed out.")
    except Exception as ex:
        print("LAN Error:", ex)
    finally:
        s.close()
        print("Socket closed.")

#############################
# USB HID Example (DLL)
#############################

import ctypes
import os

def init_and_read_idn_usb(device_index=0, lock_front_panel=False):
    dll_folder = os.path.abspath(".")
    hid_dll_path = os.path.join(dll_folder, "SLABHIDtoUART.dll")
    hid_dll = ctypes.WinDLL(hid_dll_path)

    HidUart_GetNumDevices = hid_dll.HidUart_GetNumDevices
    HidUart_Open = hid_dll.HidUart_Open
    HidUart_Close = hid_dll.HidUart_Close
    HidUart_Read = hid_dll.HidUart_Read
    HidUart_Write = hid_dll.HidUart_Write
    HidUart_SetUartConfig = hid_dll.HidUart_SetUartConfig
    HidUart_SetTimeouts = hid_dll.HidUart_SetTimeouts
    HidUart_FlushBuffers = hid_dll.HidUart_FlushBuffers

    HID_UART_SUCCESS = 0
    HID_UART_DEVICE = ctypes.c_void_p
    VID = 0x10C4
    PID = 0x8805

    def check_success(ret, msg=""):
        if ret != HID_UART_SUCCESS:
            raise IOError(f"{msg} (errcode={ret})")

    num = ctypes.c_ulong(0)
    ret = HidUart_GetNumDevices(ctypes.byref(num), VID, PID)
    check_success(ret, "HidUart_GetNumDevices failed")

    if num.value == 0:
        raise IOError("No M2000 USB HID device found")

    dev_handle = HID_UART_DEVICE()
    ret = HidUart_Open(ctypes.byref(dev_handle), device_index, VID, PID)
    check_success(ret, "HidUart_Open failed")

    ret = HidUart_SetUartConfig(
        dev_handle,
        ctypes.c_ulong(115200),
        ctypes.c_ubyte(8),
        ctypes.c_ubyte(0),
        ctypes.c_ubyte(0),
        ctypes.c_ubyte(2)
    )
    check_success(ret, "HidUart_SetUartConfig failed")

    ret = HidUart_SetTimeouts(dev_handle, 500, 500)
    check_success(ret, "HidUart_SetTimeouts failed")

    ret = HidUart_FlushBuffers(dev_handle, True, True)
    check_success(ret, "HidUart_FlushBuffers failed")

    def usb_write(cmd_string):
        out = (cmd_string + "\n").encode('ascii')
        written = ctypes.c_ulong(0)
        ret_write = HidUart_Write(dev_handle, out, len(out), ctypes.byref(written))
        check_success(ret_write, "HidUart_Write failed")
        if written.value != len(out):
            raise IOError("Not all bytes written")

    def usb_read_line(timeout_s=1.0):
        start = time.time()
        line_buf = bytearray()
        chunk_size = 256
        tmp = (ctypes.c_ubyte * chunk_size)()
        bytes_read = ctypes.c_ulong(0)

        while True:
            if (time.time() - start) > timeout_s:
                raise TimeoutError("usb_read_line timed out")

            ret_read = HidUart_Read(dev_handle, tmp, chunk_size, ctypes.byref(bytes_read))
            check_success(ret_read, "HidUart_Read failed")

            if bytes_read.value:
                for i in range(bytes_read.value):
                    line_buf.append(tmp[i])
                if b'\n' in line_buf:
                    break
            else:
                time.sleep(0.01)

        return line_buf.split(b'\n', 1)[0].decode('ascii').strip()

    try:
        usb_write("*CLS")
        time.sleep(0.1)

        if lock_front_panel:
            usb_write("LOCKOUT")
            time.sleep(0.1)

        usb_write("*IDN?")
        idn_line = usb_read_line(timeout_s=1.0)
        print("Raw *IDN? response:", idn_line)

        fields = idn_line.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

    finally:
        if dev_handle:
            HidUart_Close(dev_handle)
            print("Closed USB HID device.")

# Example usage:
if __name__ == "__main__":
    # Choose ONE of these based on your connection type:
    
    # For RS232/USB-Serial:
    # init_and_read_idn_rs232(port_name="COM4", baud_rate=115200)
    
    # For LAN:
    # init_and_read_idn_lan(ip_address="192.168.1.100", port=10733)
    
    # For USB HID:
    # init_and_read_idn_usb(device_index=0)
    
    pass  # Remove this and uncomment the method you want to use

#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) Remote Control Examples
Using RS232, LAN (TCP), or USB (HID) in Python

Demonstrates:
- Basic "init" sequence to ensure REMOTE mode
- Query instrument ID string (*IDN?)
- Optionally LOCKOUT the front panel
- Parse *IDN? response (6 fields)
- Closes the connection

Adapt as needed for your environment.
"""

import time

#############################
# 1. RS232 EXAMPLE (pyserial)
#############################
def init_and_read_idn_rs232(
    port_name="COM1", 
    baud_rate=115200, 
    lock_front_panel=False
):
    """
    Opens RS232 communication to the M2000 via pyserial,
    performs an "init" (clears interface, optionally LOCKOUT),
    requests *IDN?, and parses/prints the ID string.
    """
    import serial  # pip install pyserial

    ser = serial.Serial()
    ser.port = port_name
    ser.baudrate = baud_rate
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0         # read timeout (seconds)
    ser.rtscts   = True        # hardware flow control
    ser.dtr      = True        # must assert DTR so M2000 sees a controller
    
    try:
        ser.open()
        # 1) Clear interface
        ser.write(b"*CLS\n")
        time.sleep(0.1)

        # (Optional) 2) Lock out front panel
        if lock_front_panel:
            ser.write(b"LOCKOUT\n")
            time.sleep(0.1)

        # 3) Request ID
        ser.write(b"*IDN?\n")
        response = ser.readline().decode('ascii').strip()

        print("[RS232] Raw IDN response:", response)
        # The M2000 typically responds with up to six comma-separated fields, e.g.:
        #  Vitrek,M2000/...,12345,2,0,49
        # Let's parse them:
        fields = response.split(',')
        if len(fields) >= 1:
            manufacturer = fields[0]
        if len(fields) >= 2:
            model = fields[1]
        if len(fields) >= 3:
            serial_number = fields[2]
        if len(fields) >= 4:
            fw_major = fields[3]
        if len(fields) >= 5:
            fw_minor = fields[4]
        if len(fields) >= 6:
            fw_build = fields[5]

        print("[RS232] Parsed IDN fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Done. Optionally, if you need to revert to local:
        # ser.write(b"LOCAL\n")

    except serial.SerialException as e:
        print("RS232 Serial Error:", e)
    except Exception as e:
        print("General Exception:", e)
    finally:
        if ser.is_open:
            ser.close()


#########################
# 2. LAN (TCP/IP) EXAMPLE
#########################
def init_and_read_idn_lan(
    ip_address="192.168.1.100",
    port=10733,
    lock_front_panel=False
):
    """
    Opens a TCP socket to M2000 at IP:port,
    performs an "init" (clears interface, optionally LOCKOUT),
    requests *IDN?, and parses/prints the ID string.
    """
    import socket

    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(2.0)  # 2s timeout for connect/read
    try:
        print(f"[LAN] Connecting to {ip_address}:{port} ...")
        s.connect((ip_address, port))
        print("[LAN] Connected.")

        # 1) Clear interface
        s.sendall(b"*CLS\n")
        time.sleep(0.1)

        # (Optional) 2) Lock out front panel
        if lock_front_panel:
            s.sendall(b"LOCKOUT\n")
            time.sleep(0.1)

        # 3) Request ID
        s.sendall(b"*IDN?\n")
        
        # Read until newline
        response_buffer = []
        while True:
            chunk = s.recv(4096)
            if not chunk:
                break
            response_buffer.append(chunk)
            if b"\n" in chunk:
                break

        raw_response = b"".join(response_buffer).decode("ascii").strip()
        print("[LAN] Raw IDN response:", raw_response)

        # Parse
        fields = raw_response.split(',')
        print("[LAN] Parsed IDN fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Done. Optionally revert to local:
        # s.sendall(b"LOCAL\n")

    except socket.timeout:
        print("Socket timed out.")
    except Exception as ex:
        print("LAN error:", ex)
    finally:
        s.close()


###################################
# 3. USB (HID) EXAMPLE (ctypes + DLL)
###################################
def init_and_read_idn_usb(
    device_index=0,
    lock_front_panel=False
):
    """
    Opens USB HID connection using the Silicon Labs HID-to-UART DLL
    (SLABHIDtoUART.dll) as described in the M2000 manual.
    Performs an "init" (clears interface, optionally LOCKOUT),
    requests *IDN?, and parses the response.
    """
    import ctypes
    import os
    # Adjust as needed for your environment and the location of the DLL files
    dll_folder = os.path.abspath(".")
    hid_dll_path = os.path.join(dll_folder, "SLABHIDtoUART.dll")

    # Load the DLL
    hid_dll = ctypes.WinDLL(hid_dll_path)

    # Define function prototypes (abbreviated)
    HidUart_GetNumDevices = hid_dll.HidUart_GetNumDevices
    HidUart_Open          = hid_dll.HidUart_Open
    HidUart_Close         = hid_dll.HidUart_Close
    HidUart_Read          = hid_dll.HidUart_Read
    HidUart_Write         = hid_dll.HidUart_Write
    HidUart_SetUartConfig = hid_dll.HidUart_SetUartConfig
    HidUart_SetTimeouts   = hid_dll.HidUart_SetTimeouts
    HidUart_FlushBuffers  = hid_dll.HidUart_FlushBuffers

    HID_UART_SUCCESS = 0
    HID_UART_DEVICE = ctypes.c_void_p
    VID = 0x10C4  # 4292 in decimal
    PID = 0x8805  # 34869 in decimal

    def check_success(ret, msg=""):
        if ret != HID_UART_SUCCESS:
            raise IOError(f"{msg} (errcode={ret})")

    # 1) Find & open device
    num = ctypes.c_ulong(0)
    ret = HidUart_GetNumDevices(ctypes.byref(num), VID, PID)
    check_success(ret, "HidUart_GetNumDevices failed")

    if num.value == 0:
        raise IOError("No M2000 USB HID device found")

    dev_handle = HID_UART_DEVICE()
    ret = HidUart_Open(ctypes.byref(dev_handle), device_index, VID, PID)
    check_success(ret, "HidUart_Open failed")

    # Configure the M2000 side for 115200 8N1 w/RTS-CTS
    ret = HidUart_SetUartConfig(
        dev_handle,
        ctypes.c_ulong(115200), # Baud
        ctypes.c_ubyte(8),      # Data bits
        ctypes.c_ubyte(0),      # Parity=NONE
        ctypes.c_ubyte(0),      # Stopbits=1
        ctypes.c_ubyte(2)       # FlowControl=2 => RTS/CTS
    )
    check_success(ret, "HidUart_SetUartConfig failed")

    # Set read/write timeouts in ms
    ret = HidUart_SetTimeouts(dev_handle, 500, 500)
    check_success(ret, "HidUart_SetTimeouts failed")

    # Flush
    ret = HidUart_FlushBuffers(dev_handle, True, True)
    check_success(ret, "HidUart_FlushBuffers failed")

    # Helper for writing
    def usb_write(cmd_string):
        # M2000 expects ASCII lines, typically ending in \n
        out = (cmd_string + "\n").encode('ascii')
        written = ctypes.c_ulong(0)
        retw = HidUart_Write(dev_handle, out, len(out), ctypes.byref(written))
        check_success(retw, "HidUart_Write failed")
        if written.value != len(out):
            raise IOError("Not all bytes written")

    # Helper for reading a line
    def usb_read_line(timeout_s=1.0):
        import time
        start = time.time()
        import collections
        line_buf = bytearray()
        chunk_size = 256
        tmp = (ctypes.c_ubyte * chunk_size)()
        bytes_read = ctypes.c_ulong(0)

        while True:
            # Timeout check
            if (time.time() - start) > timeout_s:
                raise TimeoutError("usb_read_line timed out waiting for LF")

            retr = HidUart_Read(dev_handle, tmp, chunk_size, ctypes.byref(bytes_read))
            check_success(retr, "HidUart_Read failed")

            if bytes_read.value:
                for i in range(bytes_read.value):
                    line_buf.append(tmp[i])
                if b'\n' in line_buf:
                    break
            else:
                time.sleep(0.01)

        # Split at first newline
        return line_buf.split(b'\n', 1)[0].decode('ascii').strip()

    try:
        # 2) Clear interface
        usb_write("*CLS")
        time.sleep(0.1)

        # (Optional) 3) Lock out front panel
        if lock_front_panel:
            usb_write("LOCKOUT")
            time.sleep(0.1)

        # 4) Request ID
        usb_write("*IDN?")
        resp_line = usb_read_line(timeout_s=1.0)
        print("[USB] Raw IDN response:", resp_line)

        fields = resp_line.split(',')
        print("[USB] Parsed IDN fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Done. Optionally revert to local:
        # usb_write("LOCAL")

    except Exception as ex:
        print("USB HID Error:", ex)
    finally:
        # Close the device
        if dev_handle:
            HidUart_Close(dev_handle)


###########################
# Main demo of all 3 in one
###########################
if __name__ == "__main__":

    print("==== RS232 Demo ====")
    init_and_read_idn_rs232(
        port_name="COM3",
        baud_rate=115200,
        lock_front_panel=False
    )

    print("\n==== LAN Demo ====")
    init_and_read_idn_lan(
        ip_address="192.168.1.100", 
        port=10733,
        lock_front_panel=False
    )

    print("\n==== USB (HID) Demo ====")
    init_and_read_idn_usb(
        device_index=0,
        lock_front_panel=False
    )

    print("\nDone. (Adjust code as needed for your environment.)")

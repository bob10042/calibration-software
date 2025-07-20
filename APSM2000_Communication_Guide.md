# APSM2000 (M2000 Series) Python Communications Guide

This guide provides Python examples and instructions for communicating with an APSM2000 (M2000 Series) using:
- RS232 (including USB-to-RS232 adapters)
- LAN (TCP/IP)
- USB HID

## Table of Contents
1. [RS232 Communication](#rs232-communication)
2. [LAN Communication](#lan-communication) 
3. [USB HID Communication](#usb-hid-communication)
4. [Common Commands](#common-commands)
5. [Notes & Tips](#notes--tips)

## Important Notes

- The APSM2000 automatically enters REMOTE mode when it receives any valid command
- Use LOCKOUT to disable front panel control
- Use LOCAL to return to front panel control
- Commands typically end with \n (newline)
- Responses are usually comma-separated values

## RS232 Communication

### Hardware Setup
- If using PC serial port: Use null-modem (crossover) cable with full handshake lines
- If using USB: Install USB-to-RS232 adapter drivers to get a COM port
- Configure APSM2000 serial settings to match your code

### Python Code (RS232)

```python
import time
import serial  # pip install pyserial

def init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False):
    """
    Example of communicating with the APSM2000 (M2000 Series) via RS232 (or USB–RS232 adapter).
    1) Opens the serial port (assigned by OS).
    2) Sends *CLS to clear interface.
    3) (Optionally) LOCKOUT to prevent front-panel changes.
    4) Requests *IDN?.
    5) Prints and parses the returned ID string.
    6) Closes the serial port.
    """
    ser = serial.Serial()
    ser.port = port_name
    ser.baudrate = baud_rate

    # 8 data bits, no parity, 1 stop bit
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0  # seconds read timeout

    # If APSM2000 is configured for hardware flow control:
    ser.rtscts = True
    # Assert DTR so APSM2000 recognizes a controller is present:
    ser.dtr = True

    try:
        ser.open()
        print(f"Opened {port_name} at {baud_rate} baud.")

        # Clear interface
        ser.write(b"*CLS\n")
        time.sleep(0.1)

        # (Optional) Lock out front panel
        if lock_front_panel:
            ser.write(b"LOCKOUT\n")
            time.sleep(0.1)

        # Request ID
        ser.write(b"*IDN?\n")
        response = ser.readline().decode('ascii', errors='replace').strip()
        print("Raw *IDN? response:", response)

        # Parse the comma-separated fields
        fields = response.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Example of reading 3-channel voltages:
        # ser.write(b"READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3\n")
        # volt_resp = ser.readline().decode('ascii').strip()
        # print("3-Ch Voltage response:", volt_resp)

        # Revert to local if you want to unlock panel:
        # ser.write(b"LOCAL\n")

    except serial.SerialException as e:
        print("Serial error:", e)
    except Exception as e:
        print("General exception:", e)
    finally:
        if ser.is_open:
            ser.close()
            print(f"Closed {port_name}")

if __name__ == "__main__":
    init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False)
```

### Running RS232 Code
1. Install pyserial: `pip install pyserial`
2. Update port_name to match your system (e.g. "COM4" or "/dev/ttyUSB0")
3. Ensure APSM2000 serial settings match your code
4. Run the script

## LAN Communication

### Network Setup
- Connect Ethernet cable to APSM2000
- Ensure valid IP address (DHCP or manual)
- Default port is 10733
- Only one connection allowed at a time

### Python Code (LAN)

```python
import socket
import time

def init_and_read_idn_lan(ip_address="192.168.1.100", port=10733, lock_front_panel=False):
    """
    Opens a TCP socket to the APSM2000 at IP:port.
    1) Clears interface.
    2) (Optionally) LOCKOUT the front panel.
    3) Requests *IDN?.
    4) Prints ID string.
    5) Closes connection.
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(2.0)  # 2 second timeout
    try:
        print(f"Connecting to {ip_address}:{port} ...")
        s.connect((ip_address, port))
        print("Connected.")

        # Clear interface
        s.sendall(b"*CLS\n")
        time.sleep(0.1)

        if lock_front_panel:
            s.sendall(b"LOCKOUT\n")
            time.sleep(0.1)

        # Request ID
        s.sendall(b"*IDN?\n")
        
        response_buffer = []
        while True:
            chunk = s.recv(4096)
            if not chunk:
                break
            response_buffer.append(chunk)
            if b"\n" in chunk:
                # Found the newline terminator
                break

        raw_response = b"".join(response_buffer).decode('ascii').strip()
        print("Raw *IDN? response:", raw_response)

        fields = raw_response.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Example: read 3-channel voltages
        # s.sendall(b"READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3\n")
        # resp_data = s.recv(4096).decode('ascii').strip()
        # print("3-ch voltage response:", resp_data)

        # Optionally revert to local:
        # s.sendall(b"LOCAL\n")

    except socket.timeout:
        print("Socket timed out.")
    except Exception as ex:
        print("LAN Error:", ex)
    finally:
        s.close()
        print("Socket closed.")

if __name__ == "__main__":
    init_and_read_idn_lan("192.168.1.100", 10733, lock_front_panel=False)
```

### Running LAN Code
1. Verify APSM2000 IP address
2. Update ip_address in script
3. Run the script

## USB HID Communication

### Hardware Setup
- Use USB A-B cable
- No special driver needed (uses HID class)
- Need Silicon Labs DLLs for communication

### Python Code (USB HID)

```python
import ctypes
import time
import os

def init_and_read_idn_usb(device_index=0, lock_front_panel=False):
    """
    Opens USB HID connection to APSM2000 using the Silicon Labs HID DLL.
    1) Clears interface.
    2) (Optionally) LOCKOUT the front panel.
    3) Requests *IDN?.
    4) Prints ID string.
    5) Closes device.
    """
    # Adjust path to the DLLs
    dll_folder = os.path.abspath(".")
    hid_dll_path = os.path.join(dll_folder, "SLABHIDtoUART.dll")

    # Load the DLL
    hid_dll = ctypes.WinDLL(hid_dll_path)

    # Common function prototypes
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
    VID = 0x10C4  # 4292
    PID = 0x8805  # 34869

    def check_success(ret, msg=""):
        if ret != HID_UART_SUCCESS:
            raise IOError(f"{msg} (errcode={ret})")

    # Find device
    num = ctypes.c_ulong(0)
    ret = HidUart_GetNumDevices(ctypes.byref(num), VID, PID)
    check_success(ret, "HidUart_GetNumDevices failed")

    if num.value == 0:
        raise IOError("No M2000 USB HID device found")

    dev_handle = HID_UART_DEVICE()
    ret = HidUart_Open(ctypes.byref(dev_handle), device_index, VID, PID)
    check_success(ret, "HidUart_Open failed")

    # Configure 115200, 8N1, RTS/CTS
    ret = HidUart_SetUartConfig(
        dev_handle,
        ctypes.c_ulong(115200),  # baud
        ctypes.c_ubyte(8),       # data bits
        ctypes.c_ubyte(0),       # parity=none
        ctypes.c_ubyte(0),       # stopbits=1
        ctypes.c_ubyte(2)        # flow control=RTS/CTS
    )
    check_success(ret, "HidUart_SetUartConfig failed")

    # Set read/write timeouts in ms
    ret = HidUart_SetTimeouts(dev_handle, 500, 500)
    check_success(ret, "HidUart_SetTimeouts failed")

    # Flush
    ret = HidUart_FlushBuffers(dev_handle, True, True)
    check_success(ret, "HidUart_FlushBuffers failed")

    # Helper write function
    def usb_write(cmd_string):
        # M2000 expects ASCII lines, end with "\n"
        out = (cmd_string + "\n").encode('ascii')
        written = ctypes.c_ulong(0)
        ret_write = HidUart_Write(dev_handle, out, len(out), ctypes.byref(written))
        check_success(ret_write, "HidUart_Write failed")
        if written.value != len(out):
            raise IOError("Not all bytes written")

    # Helper read line function
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
        # Clear interface
        usb_write("*CLS")

        # (Optional) Lock front panel
        if lock_front_panel:
            usb_write("LOCKOUT")

        # *IDN?
        usb_write("*IDN?")
        idn_line = usb_read_line(timeout_s=1.0)
        print("Raw *IDN? response:", idn_line)

        fields = idn_line.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Example: read channel voltages
        # usb_write("READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3")
        # volt_line = usb_read_line(timeout_s=1.0)
        # print("3-ch voltage response:", volt_line)

        # Optionally go local:
        # usb_write("LOCAL")

    finally:
        if dev_handle:
            HidUart_Close(dev_handle)
            print("Closed USB HID device.")

if __name__ == "__main__":
    init_and_read_idn_usb(device_index=0, lock_front_panel=False)
```

### Running USB HID Code
1. Place SLABHIDtoUART.dll in script directory
2. Adjust device_index if multiple devices
3. Run script

## Common Commands

- `*CLS` - Clear interface/errors
- `LOCKOUT` - Lock front panel
- `LOCAL` - Return to local control
- `*IDN?` - Get device info (returns comma-separated fields):
  - Manufacturer (e.g. Vitrek)
  - Model (M2000 + options)
  - Serial Number
  - Firmware Major Version
  - Firmware Minor Version
  - Firmware Build Number
- `READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3` - Read 3 channel voltages

## Notes & Tips

### RS232/USB-Serial
- Enable RTS/CTS if configured on APSM2000
- Assert DTR (ser.dtr=True)
- Use null-modem cable with all handshake lines
- Add small delays if seeing corruption

### LAN
- Only one TCP connection at a time
- Wait ~1 min after connection loss
- Default port is 10733

### USB HID
- VID = 0x10C4 (4292)
- PID = 0x8805 (34869)
- Requires Silicon Labs DLLs
- No special driver needed on Windows

### General
- Commands end with \n
- APSM2000 enters REMOTE on any valid command
- Use LOCKOUT to prevent front panel changes
- Use LOCAL to return control

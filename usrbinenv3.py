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

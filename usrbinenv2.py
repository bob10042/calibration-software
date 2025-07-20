#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) USB HID Streaming Example

- Opens a USB HID connection using the Silicon Labs HID DLL.
- Continuously queries AC/DC voltages (or any other front-panel style measurements) from 3 channels.
- Displays them at 3 decimal places in the console.
- Logs them to a CSV file with timestamps.
- Stops on Ctrl+C.

Requirements:
- 'SLABHIDtoUART.dll' in the same folder or in system PATH.
- APSM2000 configured for 115200 8N1 RTS/CTS on its USB interface.
- Python 3.x
"""

import ctypes
import time
import os
import csv
import sys


#######################################
# 1) USB HID -> M2000 Setup & Helpers
#######################################
# Constants from M2000 documentation / Silicon Labs references
VID = 0x10C4  # 4292 decimal
PID = 0x8805 # 34869 decimal

HID_UART_SUCCESS = 0
HID_UART_DEVICE  = ctypes.c_void_p

# Adjust as needed for your environment
DLL_FOLDER      = os.path.abspath(".")
DLL_FILENAME    = "SLABHIDtoUART.dll"
DLL_PATH        = os.path.join(DLL_FOLDER, DLL_FILENAME)

# The M2000 typically needs 115200, 8N1, RTS/CTS for USB->UART bridging
BAUD_RATE       = 115200
DATA_BITS       = 8
PARITY_NONE     = 0
STOP_BITS_1     = 0
FLOW_CONTROL_RTS_CTS = 2

READ_TIMEOUT_MS  = 500
WRITE_TIMEOUT_MS = 500


#############################
# 2) Load the Silicon Labs DLL
#############################
def load_silabs_dll():
    """
    Loads the SLABHIDtoUART.dll library via ctypes.WinDLL.
    Returns the handle to the DLL and the required function references.
    Raises IOError if DLL not found or a function fails to load.
    """
    # Load the library
    try:
        hid_dll = ctypes.WinDLL(DLL_PATH)
    except OSError as e:
        raise IOError(f"Could not load {DLL_PATH}: {e}")

    # Grab needed functions
    def get_func(name, argtypes=None, restype=ctypes.c_int):
        func = getattr(hid_dll, name, None)
        if not func:
            raise AttributeError(f"Function '{name}' not found in {DLL_FILENAME}")
        func.argtypes = argtypes if argtypes else []
        func.restype  = restype
        return func

    HidUart_GetNumDevices = get_func("HidUart_GetNumDevices", [
        ctypes.POINTER(ctypes.c_ulong),  # numDevices
        ctypes.c_ushort,                 # vid
        ctypes.c_ushort                  # pid
    ])

    HidUart_Open = get_func("HidUart_Open", [
        ctypes.POINTER(HID_UART_DEVICE), # device
        ctypes.c_ulong,                  # deviceIndex
        ctypes.c_ushort,                 # vid
        ctypes.c_ushort                  # pid
    ])

    HidUart_Close = get_func("HidUart_Close", [
        HID_UART_DEVICE
    ])

    HidUart_Read = get_func("HidUart_Read", [
        HID_UART_DEVICE,
        ctypes.c_void_p,
        ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_ulong)
    ])

    HidUart_Write = get_func("HidUart_Write", [
        HID_UART_DEVICE,
        ctypes.c_void_p,
        ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_ulong)
    ])

    HidUart_SetUartConfig = get_func("HidUart_SetUartConfig", [
        HID_UART_DEVICE,
        ctypes.c_ulong, # baudRate
        ctypes.c_ubyte, # dataBits
        ctypes.c_ubyte, # parity
        ctypes.c_ubyte, # stopBits
        ctypes.c_ubyte  # flowControl
    ])

    HidUart_SetTimeouts = get_func("HidUart_SetTimeouts", [
        HID_UART_DEVICE,
        ctypes.c_ulong, # readTimeout_ms
        ctypes.c_ulong  # writeTimeout_ms
    ])

    HidUart_FlushBuffers = get_func("HidUart_FlushBuffers", [
        HID_UART_DEVICE,
        ctypes.c_bool, # flushTransmit
        ctypes.c_bool  # flushReceive
    ])

    return {
        "dll": hid_dll,
        "GetNumDevices": HidUart_GetNumDevices,
        "Open": HidUart_Open,
        "Close": HidUart_Close,
        "Read": HidUart_Read,
        "Write": HidUart_Write,
        "SetUartConfig": HidUart_SetUartConfig,
        "SetTimeouts": HidUart_SetTimeouts,
        "FlushBuffers": HidUart_FlushBuffers
    }


############################
# 3) USB HID M2000 Class
############################
class APSM2000_USB:
    """
    Encapsulates the USB HID connection to the APSM2000,
    providing simple open/close, write, and readline methods.
    """

    def __init__(self, device_index=0):
        self.device_index = device_index
        self.dll_funcs    = load_silabs_dll()
        self.dev_handle   = HID_UART_DEVICE()

    def open(self):
        # 1) Count devices
        num_devices = ctypes.c_ulong(0)
        ret = self.dll_funcs["GetNumDevices"](
            ctypes.byref(num_devices),
            VID,
            PID
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"HidUart_GetNumDevices failed (err={ret})")

        if num_devices.value == 0:
            raise IOError("No APSM2000 (M2000) USB HID devices found.")

        # 2) Open device
        ret = self.dll_funcs["Open"](
            ctypes.byref(self.dev_handle),
            self.device_index,
            VID,
            PID
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"HidUart_Open failed (err={ret})")

        # 3) Configure UART bridging
        ret = self.dll_funcs["SetUartConfig"](
            self.dev_handle,
            ctypes.c_ulong(BAUD_RATE),
            ctypes.c_ubyte(DATA_BITS),
            ctypes.c_ubyte(PARITY_NONE),
            ctypes.c_ubyte(STOP_BITS_1),
            ctypes.c_ubyte(FLOW_CONTROL_RTS_CTS)
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"HidUart_SetUartConfig failed (err={ret})")

        # 4) Timeouts
        ret = self.dll_funcs["SetTimeouts"](self.dev_handle, READ_TIMEOUT_MS, WRITE_TIMEOUT_MS)
        if ret != HID_UART_SUCCESS:
            raise IOError(f"HidUart_SetTimeouts failed (err={ret})")

        # 5) Flush buffers
        ret = self.dll_funcs["FlushBuffers"](self.dev_handle, True, True)
        if ret != HID_UART_SUCCESS:
            raise IOError(f"HidUart_FlushBuffers failed (err={ret})")

        print("APSM2000 USB HID opened successfully.")

    def close(self):
        if self.dev_handle:
            self.dll_funcs["Close"](self.dev_handle)
            print("APSM2000 USB HID closed.")
            self.dev_handle = None

    def write_line(self, command_str):
        """
        Sends a single command line (ASCII) to the M2000,
        automatically appending '\n'.
        """
        if not self.dev_handle:
            raise IOError("Device not open.")

        out_bytes = (command_str + "\n").encode('ascii')
        written   = ctypes.c_ulong(0)
        ret = self.dll_funcs["Write"](
            self.dev_handle,
            out_bytes,
            len(out_bytes),
            ctypes.byref(written)
        )
        if ret != HID_UART_SUCCESS:
            raise IOError(f"HidUart_Write failed (err={ret})")
        if written.value != len(out_bytes):
            raise IOError("HidUart_Write: Incomplete write.")

    def read_line(self, timeout_sec=1.0):
        """
        Reads ASCII data until '\n' is found, or until timeout.
        Returns the line (without trailing newline).
        """
        if not self.dev_handle:
            raise IOError("Device not open.")

        start_time = time.time()
        line_buf   = bytearray()
        chunk_size = 256
        temp_array = (ctypes.c_ubyte * chunk_size)()
        bytes_read = ctypes.c_ulong(0)

        while True:
            # Check for timeout
            if (time.time() - start_time) > timeout_sec:
                raise TimeoutError("read_line timed out waiting for newline.")

            # Attempt read
            ret = self.dll_funcs["Read"](
                self.dev_handle,
                temp_array,
                chunk_size,
                ctypes.byref(bytes_read)
            )
            if ret != HID_UART_SUCCESS:
                raise IOError(f"HidUart_Read failed (err={ret})")

            if bytes_read.value > 0:
                for i in range(bytes_read.value):
                    line_buf.append(temp_array[i])
                # Check for newline
                if b'\n' in line_buf:
                    break
            else:
                time.sleep(0.01)

        # Split at first newline
        line, _, _ = line_buf.partition(b'\n')
        return line.decode('ascii', errors='replace').strip()


############################################
# 4) Main Routine: Stream & Data-Log Example
############################################
def main_stream_voltages_and_log(
    device_index=0,
    output_csv="apms2000_datalog.csv",
    poll_interval=1.0
):
    """
    1) Opens the M2000 over USB HID.
    2) Continuously queries AC/DC voltages from 3 channels (example).
    3) Prints them to the console at 3 decimal places.
    4) Logs them to a CSV file with timestamps.
    5) Stops on Ctrl+C.
    """

    # Here is an example "READ?" command to fetch AC/DC voltages from 3 channels
    # You can add more parameters if needed (e.g. amps, watts, PF, etc.)
    # For instance:
    # READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC
    # If you also want current, you'd do:
    # READ? VOLTS:CH1:ACDC, AMPS:CH1:ACDC, ... etc.

    read_command = "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC"

    # Open the connection
    m2000 = APSM2000_USB(device_index=device_index)
    m2000.open()

    # Clear interface
    m2000.write_line("*CLS")
    time.sleep(0.2)

    # Optionally lockout front panel so user can't change settings:
    # m2000.write_line("LOCKOUT")

    # We can also do *IDN? if you want ID info:
    # m2000.write_line("*IDN?")
    # idn_resp = m2000.read_line()
    # print("IDN =", idn_resp)

    # Prepare CSV file
    csvfile = open(output_csv, mode="w", newline="")
    writer  = csv.writer(csvfile)
    # Write CSV header
    writer.writerow([
        "Timestamp (s)",
        "Voltage1 (AC+DC)",
        "Voltage2 (AC+DC)",
        "Voltage3 (AC+DC)"
    ])

    print(f"Logging data to '{output_csv}'. Press Ctrl+C to stop.\n")
    print(f"{'Time(s)':>8} | {'V1':>10} | {'V2':>10} | {'V3':>10}")

    start_time = time.time()
    try:
        while True:
            # Send the READ command
            m2000.write_line(read_command)
            resp_line = m2000.read_line(timeout_sec=2.0)

            # The response is typically comma-separated, e.g.:
            #  +1.23456E+01,+2.34567E+01,+3.45678E+01
            # Convert each to float
            parts = resp_line.split(',')
            if len(parts) < 3:
                # Incomplete data?
                continue

            # Parse floats
            try:
                v1 = float(parts[0])
                v2 = float(parts[1])
                v3 = float(parts[2])
            except ValueError:
                # If parse fails, skip
                continue

            # Current time
            elapsed_s = time.time() - start_time

            # Print with 3 decimal places
            print(f"{elapsed_s:8.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")

            # Write to CSV
            writer.writerow([f"{elapsed_s:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])

            # Flush to ensure data is saved
            csvfile.flush()

            time.sleep(poll_interval)

    except KeyboardInterrupt:
        print("\nUser stopped streaming.")
    except Exception as ex:
        print("Error during streaming:", ex)
    finally:
        # Optionally revert to local if previously locked out
        # m2000.write_line("LOCAL")

        m2000.close()
        csvfile.close()
        print("Closed CSV file, closed M2000.")


######################
# 5) If Run as Script
######################
if __name__ == "__main__":
    # Example usage:
    #   python apms2000_usb_stream.py
    # will create "apms2000_datalog.csv" in the current folder
    # and stream measurements to console.

    # Adjust as necessary
    DEVICE_INDEX   = 0
    OUTPUT_CSV     = "apms2000_datalog.csv"
    POLL_INTERVAL  = 1.0  # seconds

    main_stream_voltages_and_log(
        device_index=DEVICE_INDEX,
        output_csv=OUTPUT_CSV,
        poll_interval=POLL_INTERVAL
    )

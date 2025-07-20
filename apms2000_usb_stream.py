#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) USB HID Streaming Example with Built-In Logging & Error Handling

Features:
- Opens a USB HID connection using the Silicon Labs HID DLL (SLABHIDtoUART.dll).
- Continuously queries AC/DC voltages from 3 channels, printing to console (3 decimals) and logging to a CSV file.
- Incorporates Python's `logging` module for debug/error tracking.
- Graceful exception handling.

Usage:
  python apms2000_usb_stream.py
    (Creates "apms2000_datalog.csv" with results and prints logs to console.)
"""

import ctypes
import time
import os
import csv
import sys
import logging


####################
# 1) Logging Setup
####################
def setup_logger(debug=False):
    """
    Sets up a global logger with console output.
    If debug=True, sets level to DEBUG; otherwise INFO.
    """
    logger = logging.getLogger("APSM2000_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    # Clear old handlers (if any)
    logger.handlers.clear()
    
    # Create console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    
    # Format
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch.setFormatter(formatter)
    
    logger.addHandler(ch)
    return logger

# We'll create a module-level logger object.
logger = setup_logger(debug=False)  # Default is INFO-level.


########################################
# 2) Constants for the USB HID interface
########################################
VID = 0x10C4   # 4292 decimal
PID = 0x8805  # 34869 decimal

HID_UART_SUCCESS = 0
HID_UART_DEVICE  = ctypes.c_void_p

DLL_FILENAME    = "SLABHIDtoUART.dll"
READ_TIMEOUT_MS  = 500
WRITE_TIMEOUT_MS = 500

# Typical APSM2000 bridging config:
BAUD_RATE       = 115200
DATA_BITS       = 8
PARITY_NONE     = 0
STOP_BITS_1     = 0
FLOW_CONTROL_RTS_CTS = 2


#######################################
# 3) Load the Silicon Labs HID DLL
#######################################
def load_silabs_dll(dll_folder=None):
    """
    Loads the SLABHIDtoUART.dll via ctypes.WinDLL.
    Returns a dict of references to needed functions.
    Raises IOError/AttributeError on failure.
    """
    if dll_folder is None:
        dll_folder = os.path.abspath(".")
    dll_path = os.path.join(dll_folder, DLL_FILENAME)
    logger.debug(f"Loading DLL from: {dll_path}")

    try:
        hid_dll = ctypes.WinDLL(dll_path)
        logger.debug(f"Successfully loaded: {dll_path}")
    except OSError as e:
        logger.error(f"Could not load {dll_path}: {e}")
        raise IOError(f"Could not load {dll_path}: {e}")

    def get_func(name, argtypes=None, restype=ctypes.c_int):
        func = getattr(hid_dll, name, None)
        if not func:
            msg = f"Function '{name}' not found in {DLL_FILENAME}"
            logger.error(msg)
            raise AttributeError(msg)
        func.argtypes = argtypes if argtypes else []
        func.restype  = restype
        return func

    # Gather the needed functions
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


#######################################
# 4) Class: APSM2000 over USB HID
#######################################
class APSM2000_USB:
    """
    Encapsulates the USB HID connection to the APSM2000.
    Provides open/close, write_line, read_line, plus error handling.
    """

    def __init__(self, device_index=0, dll_folder=None):
        self.device_index = device_index
        self.dll_folder   = dll_folder or os.path.abspath(".")
        self.dll_funcs    = load_silabs_dll(self.dll_folder)
        self.dev_handle   = HID_UART_DEVICE()

    def open(self):
        logger.info("Opening APSM2000 USB HID...")
        
        # Check # of devices
        num_devices = ctypes.c_ulong(0)
        ret = self.dll_funcs["GetNumDevices"](
            ctypes.byref(num_devices),
            VID,
            PID
        )
        if ret != HID_UART_SUCCESS:
            err_msg = f"HidUart_GetNumDevices failed (err={ret})"
            logger.error(err_msg)
            raise IOError(err_msg)

        if num_devices.value == 0:
            err_msg = "No APSM2000 USB HID devices found."
            logger.error(err_msg)
            raise IOError(err_msg)

        logger.debug(f"Found {num_devices.value} device(s). Attempting to open index={self.device_index}...")
        
        # Open
        ret = self.dll_funcs["Open"](
            ctypes.byref(self.dev_handle),
            self.device_index,
            VID,
            PID
        )
        if ret != HID_UART_SUCCESS:
            err_msg = f"HidUart_Open failed (err={ret})"
            logger.error(err_msg)
            raise IOError(err_msg)

        # Configure bridging
        logger.debug("Setting UART config (115200 8N1 RTS/CTS)...")
        ret = self.dll_funcs["SetUartConfig"](
            self.dev_handle,
            ctypes.c_ulong(BAUD_RATE),
            ctypes.c_ubyte(DATA_BITS),
            ctypes.c_ubyte(PARITY_NONE),
            ctypes.c_ubyte(STOP_BITS_1),
            ctypes.c_ubyte(FLOW_CONTROL_RTS_CTS)
        )
        if ret != HID_UART_SUCCESS:
            err_msg = f"HidUart_SetUartConfig failed (err={ret})"
            logger.error(err_msg)
            raise IOError(err_msg)

        # Set timeouts
        logger.debug(f"Setting timeouts: read={READ_TIMEOUT_MS} ms, write={WRITE_TIMEOUT_MS} ms")
        ret = self.dll_funcs["SetTimeouts"](
            self.dev_handle,
            READ_TIMEOUT_MS,
            WRITE_TIMEOUT_MS
        )
        if ret != HID_UART_SUCCESS:
            err_msg = f"HidUart_SetTimeouts failed (err={ret})"
            logger.error(err_msg)
            raise IOError(err_msg)

        # Flush
        logger.debug("Flushing HID buffers...")
        ret = self.dll_funcs["FlushBuffers"](
            self.dev_handle,
            True,  # flushTransmit
            True   # flushReceive
        )
        if ret != HID_UART_SUCCESS:
            err_msg = f"HidUart_FlushBuffers failed (err={ret})"
            logger.error(err_msg)
            raise IOError(err_msg)

        logger.info("APSM2000 USB HID opened successfully.")

    def close(self):
        """
        Closes the device if open.
        """
        if self.dev_handle:
            logger.info("Closing APSM2000 USB HID...")
            self.dll_funcs["Close"](self.dev_handle)
            self.dev_handle = None
            logger.info("APSM2000 USB HID closed.")

    def write_line(self, command_str):
        """
        Writes ASCII command with trailing newline.
        """
        if not self.dev_handle:
            raise IOError("Device not open.")
        msg = command_str + "\n"
        out_bytes = msg.encode('ascii')
        written   = ctypes.c_ulong(0)

        ret = self.dll_funcs["Write"](
            self.dev_handle,
            out_bytes,
            len(out_bytes),
            ctypes.byref(written)
        )
        if ret != HID_UART_SUCCESS:
            err_msg = f"HidUart_Write failed (err={ret})"
            logger.error(err_msg)
            raise IOError(err_msg)
        if written.value != len(out_bytes):
            err_msg = "HidUart_Write incomplete write."
            logger.error(err_msg)
            raise IOError(err_msg)

        logger.debug(f"Sent: {command_str}")

    def read_line(self, timeout_sec=1.0):
        """
        Reads ASCII data until newline or timeout.
        Returns the line (str), without trailing newline.
        """
        if not self.dev_handle:
            raise IOError("Device not open.")

        start_time = time.time()
        line_buf   = bytearray()
        chunk_size = 256
        temp_array = (ctypes.c_ubyte * chunk_size)()
        bytes_read = ctypes.c_ulong(0)

        while True:
            # Check for time out
            if (time.time() - start_time) > timeout_sec:
                err_msg = "read_line timed out waiting for newline."
                logger.error(err_msg)
                raise TimeoutError(err_msg)

            # Attempt read
            ret = self.dll_funcs["Read"](
                self.dev_handle,
                temp_array,
                chunk_size,
                ctypes.byref(bytes_read)
            )
            if ret != HID_UART_SUCCESS:
                err_msg = f"HidUart_Read failed (err={ret})"
                logger.error(err_msg)
                raise IOError(err_msg)

            if bytes_read.value > 0:
                for i in range(bytes_read.value):
                    line_buf.append(temp_array[i])
                if b'\n' in line_buf:
                    break
            else:
                time.sleep(0.01)

        # Split at first newline
        line, _, _ = line_buf.partition(b'\n')
        out_line = line.decode('ascii', errors='replace').strip()
        logger.debug(f"Recv: {out_line}")
        return out_line


###########################################################
# 5) Main function: Stream and Data-Log with Error Handling
###########################################################
def stream_voltages_and_log(
    device_index=0,
    output_csv="apms2000_datalog.csv",
    poll_interval=1.0,
    debug=False
):
    """
    1) Opens the M2000 over USB HID.
    2) Continuously queries AC/DC voltages from 3 channels.
    3) Prints them to the console at 3 decimal places.
    4) Logs them to a CSV file with timestamps.
    5) Includes robust error handling and debug logs if debug=True.
    6) Stops on Ctrl+C.
    """

    # Adjust global logger to desired level
    global logger
    logger = setup_logger(debug=debug)

    logger.info("=== Starting APSM2000 USB streaming with built-in logging & error handling ===")
    logger.info(f"Output CSV: {output_csv}")
    logger.info(f"Poll Interval: {poll_interval} s")
    logger.info(f"Debug Mode: {debug}")

    # Example READ? command:
    read_command = "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC"
    
    # Create M2000 object
    m2000 = APSM2000_USB(device_index=device_index)

    try:
        # Open device
        m2000.open()

        # Clear interface, flush errors
        m2000.write_line("*CLS")
        time.sleep(0.2)

        # Optional: lock front panel
        # m2000.write_line("LOCKOUT")

        # Example: query ID
        # m2000.write_line("*IDN?")
        # idn_line = m2000.read_line(timeout_sec=2.0)
        # logger.info(f"IDN Response = {idn_line}")

        # Prepare CSV
        with open(output_csv, mode="w", newline="") as csvfile:
            writer = csv.writer(csvfile)
            # Write header
            writer.writerow(["Timestamp (s)", "Voltage1 (AC+DC)", "Voltage2 (AC+DC)", "Voltage3 (AC+DC)"])
            logger.info(f"CSV logging started: {output_csv}")

            start_time = time.time()
            logger.info("Press Ctrl+C to stop streaming...")

            while True:
                try:
                    # Send read command
                    m2000.write_line(read_command)
                    response_line = m2000.read_line(timeout_sec=2.0)

                    # Split fields
                    parts = response_line.split(',')
                    if len(parts) < 3:
                        # Possibly incomplete response
                        logger.warning(f"Incomplete data from M2000: {response_line}")
                        continue

                    # Parse floats
                    try:
                        v1 = float(parts[0])
                        v2 = float(parts[1])
                        v3 = float(parts[2])
                    except ValueError as ex:
                        logger.warning(f"Could not parse floats from: {response_line} => {ex}")
                        continue

                    elapsed_s = time.time() - start_time

                    # Print to console
                    # Format to 3 decimals
                    print(f"{elapsed_s:8.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")

                    # Log to CSV
                    writer.writerow([f"{elapsed_s:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])
                    csvfile.flush()

                    time.sleep(poll_interval)

                except TimeoutError as tex:
                    logger.error(f"Timeout reading data: {tex}")
                except Exception as genex:
                    logger.error(f"Unexpected error in streaming loop: {genex}", exc_info=True)

    except KeyboardInterrupt:
        logger.info("User pressed Ctrl+C. Stopping streaming.")
    except Exception as e:
        logger.error(f"Critical error: {e}", exc_info=True)
    finally:
        # Optionally revert to local
        # m2000.write_line("LOCAL")

        # Close device
        m2000.close()
        logger.info("=== APSM2000 USB streaming stopped. ===")


######################
# 6) If run as script
######################
if __name__ == "__main__":
    # Default arguments
    DEVICE_INDEX  = 0
    OUTPUT_CSV    = "apms2000_datalog.csv"
    POLL_INTERVAL = 1.0

    # If you want debug logs, change to True
    DEBUG_MODE    = False

    # Start streaming
    stream_voltages_and_log(
        device_index=DEVICE_INDEX,
        output_csv=OUTPUT_CSV,
        poll_interval=POLL_INTERVAL,
        debug=DEBUG_MODE
    )
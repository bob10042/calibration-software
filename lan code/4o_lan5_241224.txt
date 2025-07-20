#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) LAN Script with Full Initialization, IDN?, Remote Lock,
Auto-Detect of 3-Channel Read, and Data Logging

1) Connect via TCP to IP:port.
2) *CLS to clear.
3) *IDN? to get ID string (logs it).
4) LOCKOUT (optional) to ensure front panel is locked => definitely REMOTE.
5) Tries multiple READ? variants for CH1..CH3 until one works (no timeout, valid floats).
6) Streams data to console & CSV until Ctrl+C.
7) LOCAL (optional) to unlock, then closes the socket.
"""

import socket
import time
import sys
import csv
import logging
from datetime import datetime

#########################
# 1) Logging Setup
#########################
def setup_logger(debug=False):
    """
    Sets up a global logger that prints to console.
    If debug=True, logs at DEBUG level; otherwise INFO level.
    """
    logger = logging.getLogger("APSM2000_LAN_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    # Remove old handlers (if any)
    if logger.hasHandlers():
        logger.handlers.clear()

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger

###############################
# 2) APSM2000 LAN Helper Class
###############################
class APSM2000_LAN:
    """
    Encapsulates a TCP socket connection to the APSM2000 over LAN.
    Provides open/close, write_line, read_line methods with logging & error handling.
    """

    def __init__(self, ip_address="192.168.15.100", port=10733, debug=False):
        self.ip_address = ip_address
        self.port       = port
        self.sock       = None
        self.logger     = setup_logger(debug=debug)

    def open(self, timeout=5.0):
        """
        Opens a TCP socket to the APSM2000 at ip_address:port.
        """
        self.logger.info(f"Connecting to APSM2000 at {self.ip_address}:{self.port}...")
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.settimeout(timeout)
        try:
            self.sock.connect((self.ip_address, self.port))
            self.logger.info("Successfully connected.")
        except socket.timeout:
            self.logger.error(f"Socket timed out attempting connection to {self.ip_address}:{self.port}")
            raise
        except Exception as ex:
            self.logger.error(f"Error connecting to {self.ip_address}:{self.port}: {ex}")
            raise

    def close(self):
        """
        Closes the TCP connection.
        """
        if self.sock:
            self.logger.info("Closing socket connection.")
            try:
                self.sock.close()
            except Exception as ex:
                self.logger.warning(f"Error closing socket: {ex}")
            self.sock = None

    def write_line(self, command_str):
        """
        Sends a single ASCII line with trailing newline over the socket.
        """
        if not self.sock:
            raise IOError("Socket is not open.")
        msg = (command_str + "\n").encode("ascii")
        self.logger.debug(f"Sending: {command_str}")
        try:
            self.sock.sendall(msg)
        except Exception as ex:
            self.logger.error(f"Error sending data: {ex}")
            raise

    def read_line(self, timeout_sec=2.0):
        """
        Reads until newline or until timeout.
        Returns the line as a string, without the trailing newline.
        """
        if not self.sock:
            raise IOError("Socket is not open.")

        end_time = time.time() + timeout_sec
        line_buf = bytearray()

        while True:
            if time.time() > end_time:
                raise TimeoutError("read_line timed out waiting for data.")

            try:
                chunk = self.sock.recv(256)
            except socket.timeout:
                # no data yet
                time.sleep(0.01)
                continue
            except Exception as ex:
                self.logger.error(f"Error receiving data: {ex}")
                raise

            if not chunk:
                # Connection closed or no data
                self.logger.warning("Socket recv returned empty data (connection closed?).")
                break

            line_buf.extend(chunk)
            if b'\n' in chunk:
                break

        # Split at first newline
        line, _, _ = line_buf.partition(b'\n')
        out_line = line.decode('ascii', errors='replace').strip()
        self.logger.debug(f"Received: {out_line}")
        return out_line

##############################
# 3) Main Logic
##############################
def run_lan_script(
    ip_address="192.168.15.100",
    port=10733,
    output_csv="apsm2000_lan_datalog.csv",
    poll_interval=1.0,
    debug=False
):
    """
    1) Connect to APSM2000.
    2) *CLS
    3) *IDN? => parse it
    4) LOCKOUT => ensure front panel is locked
    5) Try 3 command variants for reading CH1..CH3
       - READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC
       - READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3
       - READ? VOLTS:CH1:AC, VOLTS:CH2:AC, VOLTS:CH3:AC
    6) Once success, loop to read in CSV until Ctrl+C
    7) LOCAL => unlock front panel, close
    """

    logger = setup_logger(debug=debug)
    logger.info("=== Starting APSM2000 LAN full init & auto-detect for 3 channels ===")
    logger.info(f"IP : {ip_address}:{port}")
    logger.info(f"CSV: {output_csv}")
    logger.info(f"Poll Interval: {poll_interval}s")
    logger.info(f"Debug Mode: {debug}")

    m2000 = APSM2000_LAN(ip_address=ip_address, port=port, debug=debug)

    # Available command variants:
    read_variants = [
        "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC",
        "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3",
        "READ? VOLTS:CH1:AC, VOLTS:CH2:AC, VOLTS:CH3:AC"
    ]

    working_command = None

    try:
        # 1) Open
        m2000.open(timeout=5.0)

        # 2) Clear interface
        m2000.write_line("*CLS")
        time.sleep(0.2)

        # 3) IDN?
        m2000.write_line("*IDN?")
        time.sleep(0.1)
        try:
            idn_resp = m2000.read_line(timeout_sec=2.0)
            logger.info(f"IDN Response: {idn_resp}")
        except TimeoutError:
            logger.warning("No IDN response (timeout). Possibly the M2000 doesn't respond to *IDN?")

        # 4) LOCKOUT (optional)
        m2000.write_line("LOCKOUT")
        time.sleep(0.2)

        # 5) Try the read variants
        for variant in read_variants:
            logger.info(f"Trying command variant: {variant}")
            try:
                m2000.write_line(variant)
                resp_line = m2000.read_line(timeout_sec=2.0)
                parts = resp_line.split(',')
                if len(parts) == 3:
                    # parse floats
                    test_floats = [float(x) for x in parts]
                    # if we got here => success
                    logger.info(f"Success with '{variant}' => example = {test_floats}")
                    working_command = variant
                    break
                else:
                    logger.warning(f"Response doesn't have 3 fields: '{resp_line}'")
            except (TimeoutError, ValueError) as e:
                logger.warning(f"Variant '{variant}' failed: {e}")
            except Exception as ex:
                logger.error(f"Error with variant '{variant}': {ex}", exc_info=True)

        if not working_command:
            logger.error("No read command variant worked. Exiting.")
            return  # stop the script here

        # 6) Now loop reading with the working_command
        with open(output_csv, mode='w', newline='') as cf:
            writer = csv.writer(cf)
            # Header
            writer.writerow(["Timestamp", "Elapsed_s", "CH1_Voltage", "CH2_Voltage", "CH3_Voltage"])

            logger.info("Press Ctrl+C to stop streaming...\n")
            print(f"\n{'Elapsed(s)':>10} | {'CH1(V)':>10} | {'CH2(V)':>10} | {'CH3(V)':>10}")
            print("-" * 46)

            start_time = time.time()

            while True:
                try:
                    m2000.write_line(working_command)
                    resp_line = m2000.read_line(timeout_sec=2.0)
                    parts = resp_line.split(',')
                    if len(parts) < 3:
                        logger.error(f"Incomplete response: '{resp_line}'")
                        continue
                    try:
                        v1 = float(parts[0])
                        v2 = float(parts[1])
                        v3 = float(parts[2])
                    except ValueError:
                        logger.error(f"Could not parse floats from '{resp_line}'")
                        continue

                    elapsed_s = time.time() - start_time
                    now_str   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Print
                    print(f"{elapsed_s:10.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")

                    # CSV
                    writer.writerow([now_str, f"{elapsed_s:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])
                    cf.flush()

                    time.sleep(poll_interval)

                except TimeoutError:
                    logger.error("Timeout reading data (timed out).")
                except KeyboardInterrupt:
                    logger.info("User pressed Ctrl+C. Stopping streaming.")
                    break
                except Exception as genex:
                    logger.error(f"Unexpected error in streaming loop: {genex}", exc_info=True)

    except KeyboardInterrupt:
        logger.info("User pressed Ctrl+C during init.")
    except Exception as top_ex:
        logger.error(f"Critical error: {top_ex}", exc_info=True)
    finally:
        # 7) Optionally revert to local
        try:
            m2000.write_line("LOCAL")
            time.sleep(0.1)
        except:
            pass

        m2000.close()
        logger.info("=== APSM2000 LAN script finished ===")


##############################
# 4) If run as main script
##############################
if __name__ == "__main__":
    IP_ADDRESS  = "192.168.15.100"   # Adjust as needed
    PORT        = 10733             # M2000's default LAN port
    OUTPUT_CSV  = "apsm2000_lan_datalog.csv"
    POLL_SECS   = 1.0
    DEBUG_MODE  = False  # Set True for more debug logs

    run_lan_script(
        ip_address=IP_ADDRESS,
        port=PORT,
        output_csv=OUTPUT_CSV,
        poll_interval=POLL_SECS,
        debug=DEBUG_MODE
    )

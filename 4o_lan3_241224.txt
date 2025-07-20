#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) LAN Streaming Example (3 Channels)

- Connects via TCP/IP to the APSM2000 at a specified IP and port (default 10733).
- Continuously queries AC/DC voltages from 3 channels (CH1, CH2, CH3).
- Displays them in the console with 3 decimal places.
- Logs them in a CSV file with timestamps.
- Includes error handling and Python's logging for easier debugging.

Press Ctrl+C to stop streaming.

If the script times out when reading voltages, try:
 - Using 'READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3' (dropping ':ACDC'),
 - Or verifying you indeed have 3 channels physically installed,
 - Or verifying the M2000 is in a valid mode (MODE? => 1 for multi-VPA).
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
    If debug=True, logs at DEBUG level; otherwise INFO.
    """
    logger = logging.getLogger("APSM2000_LAN_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    # Remove old handlers
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
    Provides open/close, write_line, read_line methods with basic error handling.
    """

    def __init__(self, ip_address="192.168.15.100", port=10733, debug=False):
        self.ip_address = ip_address
        self.port       = port
        self.sock       = None
        # If debug=True, configure logger to DEBUG
        self.logger     = setup_logger(debug=debug)

    def open(self, timeout=2.0):
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
        Returns the line as a string, without trailing newline.
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
                # no data yet, small pause?
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
# 3) Main Streaming Function
##############################
def stream_three_channels_lan(
    ip_address="192.168.15.100",
    port=10733,
    output_csv="apsm2000_lan_datalog.csv",
    poll_interval=1.0,
    debug=False
):
    """
    1) Connect to the APSM2000 at ip_address:port.
    2) Clears interface (*CLS).
    3) Continuously queries the AC/DC voltages for CH1, CH2, CH3.
    4) Logs to CSV, prints 3 decimal places to console.
    5) On Ctrl+C, stops gracefully.

    If you only have 2 channels physically installed, adapt the read command accordingly.
    If your firmware doesn't recognize :ACDC, try removing it or using :AC.
    """

    logger = setup_logger(debug=debug)
    logger.info(f"=== Starting APSM2000 LAN 3-channel streaming ===")
    logger.info(f"Target IP : {ip_address}:{port}")
    logger.info(f"Output CSV: {output_csv}")
    logger.info(f"Poll Interval: {poll_interval} s")
    logger.info(f"Debug Mode: {debug}")

    # The command to read 3 channels (AC+DC):
    read_command = "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC"
    # If :ACDC times out, try "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3"

    m2000 = APSM2000_LAN(ip_address=ip_address, port=port, debug=debug)

    try:
        # 1) Open connection
        m2000.open(timeout=5.0)

        # 2) Clear interface
        m2000.write_line("*CLS")
        time.sleep(0.2)

        # Optionally lock out front panel
        # m2000.write_line("LOCKOUT")
        # time.sleep(0.2)

        # Create CSV
        with open(output_csv, mode='w', newline='') as cf:
            writer = csv.writer(cf)
            writer.writerow(["Timestamp", "Elapsed_s", "CH1_Voltage", "CH2_Voltage", "CH3_Voltage"])

            logger.info("\nPress Ctrl+C to stop streaming...\n")
            print(f"\n{'Elapsed(s)':>10} | {'CH1(V)':>10} | {'CH2(V)':>10} | {'CH3(V)':>10}")
            print("-" * 46)

            start_time = time.time()

            while True:
                try:
                    # 3) Send read command
                    m2000.write_line(read_command)

                    # 4) Read response
                    resp_line = m2000.read_line(timeout_sec=2.0)  # Wait for one line
                    # Example response: +1.23456E+01,+2.34567E+01,+3.45678E+01

                    parts = resp_line.split(',')
                    if len(parts) < 3:
                        logger.error(f"Incomplete response for 3 channels: {resp_line}")
                        continue

                    try:
                        v1 = float(parts[0])
                        v2 = float(parts[1])
                        v3 = float(parts[2])
                    except ValueError:
                        logger.error(f"Could not parse floats from {resp_line}")
                        continue

                    elapsed_s = time.time() - start_time
                    now_str   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Print to console with 3 decimals
                    print(f"{elapsed_s:10.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")

                    # Write to CSV
                    writer.writerow([now_str, f"{elapsed_s:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])
                    cf.flush()

                    time.sleep(poll_interval)

                except TimeoutError:
                    logger.error("Timeout during streaming: timed out")
                except Exception as ex:
                    logger.error(f"Unexpected error in streaming loop: {ex}", exc_info=True)

    except KeyboardInterrupt:
        logger.info("User pressed Ctrl+C. Stopping streaming.")
    except Exception as e:
        logger.error(f"Critical error: {e}", exc_info=True)
    finally:
        # Optionally revert to local
        # m2000.write_line("LOCAL")

        m2000.close()
        logger.info("=== APSM2000 LAN streaming stopped ===")


##############################
# 4) If run as main script
##############################
if __name__ == "__main__":
    # Adjust IP/port as needed
    IP_ADDRESS  = "192.168.15.100"
    PORT        = 10733
    OUTPUT_CSV  = "apsm2000_lan_datalog.csv"
    POLL_SECS   = 1.0
    DEBUG_MODE  = False  # Set True to see more debug details

    stream_three_channels_lan(
        ip_address=IP_ADDRESS,
        port=PORT,
        output_csv=OUTPUT_CSV,
        poll_interval=POLL_SECS,
        debug=DEBUG_MODE
    )

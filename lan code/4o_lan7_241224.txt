#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) LAN Full Script:
1) Connect & *CLS
2) *IDN?
3) Force MODE=1 & SAVECONFIG
4) CHNL? 1..3 (optional, if firmware supports)
5) LOCKOUT => ensure remote mode locked
6) Auto-detect 3-channel read among:
   a) READ? VOLTS:CH1:AC, VOLTS:CH2:AC, VOLTS:CH3:AC
   b) READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3
   c) READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC
7) Stream voltages to console & CSV
8) LOCAL => unlock, close

Press Ctrl+C to stop streaming.
"""

import socket
import time
import sys
import csv
import logging
from datetime import datetime

######################
# 1) Logging Setup
######################
def setup_logger(debug=False):
    logger = logging.getLogger("APSM2000_LAN_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    # Remove existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    return logger


###############################
# 2) APSM2000 LAN Helper Class
###############################
class APSM2000_LAN:
    """
    Encapsulates a TCP socket connection to the APS M2000 analyzer,
    providing write_line() & read_line() with basic error handling/logging.
    """

    def __init__(self, ip="192.168.15.100", port=10733, debug=False):
        self.ip = ip
        self.port = port
        self.sock = None
        self.logger = setup_logger(debug=debug)

    def open(self, timeout=5.0):
        self.logger.info(f"Connecting to APSM2000 at {self.ip}:{self.port}...")
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.settimeout(timeout)
        try:
            self.sock.connect((self.ip, self.port))
            self.logger.info("Successfully connected.")
        except socket.timeout:
            self.logger.error(f"Connection to {self.ip}:{self.port} timed out.")
            raise
        except Exception as ex:
            self.logger.error(f"Error connecting to {self.ip}:{self.port}: {ex}")
            raise

    def close(self):
        if self.sock:
            self.logger.info("Closing socket connection.")
            try:
                self.sock.close()
            except Exception as ex:
                self.logger.warning(f"Error during socket close: {ex}")
            self.sock = None

    def write_line(self, command):
        if not self.sock:
            raise IOError("Socket is not open.")
        data = (command + "\n").encode("ascii")
        self.logger.debug(f"Sending: {command}")
        try:
            self.sock.sendall(data)
        except Exception as ex:
            self.logger.error(f"Error sending data: {ex}")
            raise

    def read_line(self, timeout_s=2.0):
        if not self.sock:
            raise IOError("Socket is not open.")

        end_time = time.time() + timeout_s
        line_buf = bytearray()

        while True:
            if time.time() > end_time:
                raise TimeoutError("read_line timed out waiting for data.")

            try:
                chunk = self.sock.recv(256)
            except socket.timeout:
                time.sleep(0.01)
                continue
            except Exception as ex:
                self.logger.error(f"Recv error: {ex}")
                raise

            if not chunk:
                self.logger.warning("Socket recv returned empty data.")
                break

            line_buf.extend(chunk)
            if b'\n' in chunk:
                break

        # Split at first newline
        line, _, _ = line_buf.partition(b'\n')
        msg = line.decode('ascii', errors='replace').strip()
        self.logger.debug(f"Received: {msg}")
        return msg


##############################
# 3) Main Script Function
##############################
def run_lan_script(
    ip="192.168.15.100",
    port=10733,
    csv_file="apsm2000_lan_datalog.csv",
    poll_interval=1.0,
    debug=False
):
    """
    Complete flow:
    1) Connect, *CLS
    2) *IDN?
    3) MODE,1 + SAVECONFIG
    4) CHNL? 1..3
    5) LOCKOUT
    6) Auto-detect read among:
       a) READ? VOLTS:CH1:AC, VOLTS:CH2:AC, VOLTS:CH3:AC
       b) READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3
       c) READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC
    7) Stream to CSV
    8) LOCAL + close
    """
    logger = setup_logger(debug=debug)
    logger.info("=== APS M2000 LAN full script with AC coupling attempt ===")
    logger.info(f"IP : {ip}:{port}")
    logger.info(f"CSV: {csv_file}")
    logger.info(f"Poll Interval: {poll_interval}s")
    logger.info(f"Debug: {debug}")

    # We suspect "AC" is correct subfield based on the photos, so we test it first:
    read_variants = [
        "READ? VOLTS:CH1:AC, VOLTS:CH2:AC, VOLTS:CH3:AC",
        "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3",
        "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC",
    ]
    working_cmd = None

    m2000 = APSM2000_LAN(ip=ip, port=port, debug=debug)

    try:
        # 1) Connect
        m2000.open()

        # 2) Clear interface
        m2000.write_line("*CLS")
        time.sleep(0.2)

        # 3) IDN?
        m2000.write_line("*IDN?")
        time.sleep(0.1)
        try:
            idn_resp = m2000.read_line(timeout_s=2.0)
            logger.info(f"IDN: {idn_resp}")
        except TimeoutError:
            logger.warning("No IDN response (timeout).")

        # Force Multi-VPA
        m2000.write_line("MODE,1")
        time.sleep(0.2)
        m2000.write_line("SAVECONFIG")
        time.sleep(0.5)

        # 4) Query channels
        for c in [1, 2, 3]:
            cmd_ch = f"CHNL? {c}"
            logger.info(f"Querying channel {c} -> '{cmd_ch}'")
            try:
                m2000.write_line(cmd_ch)
                ch_resp = m2000.read_line(timeout_s=1.5)
                logger.info(f"CHNL? {c} => {ch_resp}")
            except TimeoutError:
                logger.warning(f"CHNL? {c} timed out.")
            time.sleep(0.1)

        # 5) Lock front panel
        m2000.write_line("LOCKOUT")
        time.sleep(0.2)

        # 6) Auto-detect read command
        for variant in read_variants:
            logger.info(f"Trying read variant: {variant}")
            try:
                m2000.write_line(variant)
                resp_line = m2000.read_line(timeout_s=2.0)
                parts = resp_line.split(',')
                if len(parts) == 3:
                    # parse floats
                    flts = [float(p) for p in parts]
                    logger.info(f"Success with '{variant}' => {flts}")
                    working_cmd = variant
                    break
                else:
                    logger.warning(f"Wrong #fields => {resp_line}")
            except (TimeoutError, ValueError) as e:
                logger.warning(f"Variant '{variant}' failed: {e}")
            except Exception as ex:
                logger.error(f"Error with '{variant}': {ex}", exc_info=True)

        if not working_cmd:
            logger.error("No read command worked. Exiting.")
            return

        # 7) Stream data to CSV
        with open(csv_file, "w", newline="") as cf:
            writer = csv.writer(cf)
            writer.writerow(["Timestamp", "Elapsed_s", "CH1_V", "CH2_V", "CH3_V"])

            logger.info("Press Ctrl+C to stop streaming...\n")
            print(f"\n{'Elapsed(s)':>10} | {'CH1(V)':>10} | {'CH2(V)':>10} | {'CH3(V)':>10}")
            print("-" * 46)

            start_time = time.time()

            while True:
                try:
                    m2000.write_line(working_cmd)
                    line = m2000.read_line(timeout_s=2.0)
                    parts = line.split(',')
                    if len(parts) < 3:
                        logger.error(f"Incomplete data => {line}")
                        continue
                    try:
                        v1 = float(parts[0])
                        v2 = float(parts[1])
                        v3 = float(parts[2])
                    except ValueError:
                        logger.error(f"Could not parse floats from '{line}'")
                        continue

                    elapsed = time.time() - start_time
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Print to console
                    print(f"{elapsed:10.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")

                    # Write CSV
                    writer.writerow([now_str, f"{elapsed:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])
                    cf.flush()

                    time.sleep(poll_interval)

                except TimeoutError:
                    logger.error("Timeout reading data.")
                except KeyboardInterrupt:
                    logger.info("Ctrl+C => Stopping streaming.")
                    break
                except Exception as loop_ex:
                    logger.error(f"Unexpected loop error: {loop_ex}", exc_info=True)

    except KeyboardInterrupt:
        logger.info("Ctrl+C during init. Aborting.")
    except Exception as main_ex:
        logger.error(f"Critical error: {main_ex}", exc_info=True)
    finally:
        # Return to local
        try:
            m2000.write_line("LOCAL")
            time.sleep(0.1)
        except:
            pass
        m2000.close()
        logger.info("=== Script finished ===")


#####################################
# If run as script:
#####################################
if __name__ == "__main__":
    IP_ADDRESS = "192.168.15.100"
    PORT       = 10733
    CSV_FILE   = "apsm2000_lan_datalog.csv"
    INTERVAL   = 1.0
    DEBUG      = False  # set True for more debug logs

    run_lan_script(
        ip=IP_ADDRESS,
        port=PORT,
        csv_file=CSV_FILE,
        poll_interval=INTERVAL,
        debug=DEBUG
    )

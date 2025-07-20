#!/usr/bin/env python3
"""
APSM2000 LAN Script:
1) Connect & *CLS
2) *IDN? -> see ID
3) Force Multi-VPA mode: MODE,1 + SAVECONFIG
4) Optional: Query channels with CHNL? 1..3
5) Auto-detect READ? variants for CH1..CH3
6) Stream data to console & CSV
7) On exit, LOCAL + close
"""

import socket
import time
import sys
import csv
import logging
from datetime import datetime

#########################
# Logging Setup
#########################
def setup_logger(debug=False):
    logger = logging.getLogger("APSM2000_LAN_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    if logger.hasHandlers():
        logger.handlers.clear()

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s","%Y-%m-%d %H:%M:%S")
    ch.setFormatter(fmt)
    logger.addHandler(ch)
    return logger


class APSM2000_LAN:
    """
    Encapsulates a TCP socket to the APSM2000 with write/read line methods.
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
            self.logger.error(f"Connection timed out to {self.ip}:{self.port}")
            raise
        except Exception as ex:
            self.logger.error(f"Error connecting: {ex}")
            raise

    def close(self):
        if self.sock:
            self.logger.info("Closing socket.")
            try:
                self.sock.close()
            except Exception as ex:
                self.logger.warning(f"Error on close: {ex}")
            self.sock = None

    def write_line(self, cmd):
        if not self.sock:
            raise IOError("Socket not open.")
        data = (cmd + "\n").encode("ascii")
        self.logger.debug(f"Sending: {cmd}")
        self.sock.sendall(data)

    def read_line(self, timeout_s=2.0):
        if not self.sock:
            raise IOError("Socket not open.")

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
            if not chunk:
                self.logger.warning("Socket recv empty (closed?).")
                break

            line_buf.extend(chunk)
            if b'\n' in chunk:
                break

        line, _, _ = line_buf.partition(b'\n')
        msg = line.decode("ascii", errors="replace").strip()
        self.logger.debug(f"Received: {msg}")
        return msg


def run_lan_script(
    ip="192.168.15.100",
    port=10733,
    csv_file="apsm2000_lan_datalog.csv",
    poll_interval=1.0,
    debug=False
):
    logger = setup_logger(debug=debug)
    logger.info("=== APSM2000 LAN Force Mode & Auto-Detect Script ===")
    logger.info(f"IP: {ip}:{port}")
    logger.info(f"CSV: {csv_file}")
    logger.info(f"Poll Interval: {poll_interval}s")
    logger.info(f"Debug: {debug}")

    m2000 = APSM2000_LAN(ip=ip, port=port, debug=debug)

    # Our read variants
    read_variants = [
        "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC",
        "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3",
        "READ? VOLTS:CH1:AC, VOLTS:CH2:AC, VOLTS:CH3:AC"
    ]

    try:
        # 1) Connect
        m2000.open()

        # 2) Clear
        m2000.write_line("*CLS")
        time.sleep(0.2)

        # 3) IDN?
        m2000.write_line("*IDN?")
        time.sleep(0.1)
        try:
            idn = m2000.read_line(timeout_s=2.0)
            logger.info(f"IDN Reply: {idn}")
        except TimeoutError:
            logger.warning("No *IDN? response")

        # Force Multi-VPA mode
        logger.info("Setting MODE,1 => multi-VPA, then SAVECONFIG")
        m2000.write_line("MODE,1")
        time.sleep(0.2)
        m2000.write_line("SAVECONFIG")
        time.sleep(0.5)

        # Optional: query channels
        for c in [1,2,3]:
            cmd_chnl = f"CHNL? {c}"
            logger.info(f"Checking channel {c} -> {cmd_chnl}")
            try:
                m2000.write_line(cmd_chnl)
                resp_chnl = m2000.read_line(timeout_s=1.5)
                logger.info(f"Channel {c} info: {resp_chnl}")
            except TimeoutError:
                logger.warning(f"CHNL? {c} timed out (maybe not supported).")
            time.sleep(0.1)

        # Now do LOCKOUT (to ensure remote mode locked)
        m2000.write_line("LOCKOUT")
        time.sleep(0.2)

        # 4) Auto-Detect read command
        working_cmd = None
        for variant in read_variants:
            logger.info(f"Trying variant: {variant}")
            try:
                m2000.write_line(variant)
                resp_line = m2000.read_line(timeout_s=2.0)
                parts = resp_line.split(',')
                if len(parts) == 3:
                    # parse floats
                    test_vals = [float(p) for p in parts]
                    logger.info(f"Success with '{variant}' => {test_vals}")
                    working_cmd = variant
                    break
                else:
                    logger.warning(f"Wrong #fields: {resp_line}")
            except (TimeoutError, ValueError) as e:
                logger.warning(f"Variant '{variant}' failed: {e}")
            except Exception as ex:
                logger.error(f"Error with variant '{variant}': {ex}", exc_info=True)

        if not working_cmd:
            logger.error("No read command worked. Exiting.")
            return

        # 5) Stream to CSV
        with open(csv_file, "w", newline="") as cf:
            writer = csv.writer(cf)
            writer.writerow(["Timestamp", "Elapsed_s", "CH1_V", "CH2_V", "CH3_V"])

            logger.info("Press Ctrl+C to stop streaming...\n")
            print(f"\n{'Elapsed(s)':>10} | {'CH1(V)':>10} | {'CH2(V)':>10} | {'CH3(V)':>10}")
            print("-"*46)

            start_time = time.time()

            while True:
                try:
                    m2000.write_line(working_cmd)
                    line = m2000.read_line(timeout_s=2.0)
                    parts = line.split(',')
                    if len(parts) < 3:
                        logger.error(f"Incomplete read: {line}")
                        continue
                    try:
                        v1 = float(parts[0])
                        v2 = float(parts[1])
                        v3 = float(parts[2])
                    except ValueError:
                        logger.error(f"Could not parse floats from '{line}'")
                        continue

                    elapsed_s = time.time() - start_time
                    now_str   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Print
                    print(f"{elapsed_s:10.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")
                    # Write CSV
                    writer.writerow([now_str, f"{elapsed_s:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])
                    cf.flush()

                    time.sleep(poll_interval)
                except TimeoutError:
                    logger.error("Timeout reading data.")
                except KeyboardInterrupt:
                    logger.info("User pressed Ctrl+C => stopping.")
                    break
                except Exception as main_ex:
                    logger.error(f"Loop error: {main_ex}", exc_info=True)

    except KeyboardInterrupt:
        logger.info("Ctrl+C during init.")
    except Exception as main_e:
        logger.error(f"Critical error: {main_e}", exc_info=True)
    finally:
        # revert to local
        try:
            m2000.write_line("LOCAL")
            time.sleep(0.1)
        except:
            pass
        m2000.close()
        logger.info("=== Finished ===")


if __name__ == "__main__":
    IP_ADDR    = "192.168.15.100"  # Adjust if needed
    PORT       = 10733
    CSV_OUTPUT = "apsm2000_lan_datalog.csv"
    POLL_TIME  = 1.0
    DEBUG_MODE = False

    run_lan_script(
        ip=IP_ADDR,
        port=PORT,
        csv_file=CSV_OUTPUT,
        poll_interval=POLL_TIME,
        debug=DEBUG_MODE
    )

#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) LAN Mega Auto-Detect Script

1) Connect & *CLS
2) *IDN? (try to see ID)
3) Force MODE=1 + SAVECONFIG
4) LOCKOUT => remote
5) Attempts 9 read variants for 3 channels, covering:
   - CHx, A x, VPAx
   - subfields :AC, :ACDC, or none
6) Once success, streams data to console & CSV
7) LOCAL + close on exit

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


class APSM2000_LAN:
    """
    Encapsulates TCP connection to APS M2000, with write_line/read_line.
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
            self.logger.error("Connection timed out.")
            raise
        except Exception as ex:
            self.logger.error(f"Error connecting: {ex}")
            raise

    def close(self):
        if self.sock:
            self.logger.info("Closing socket connection.")
            try:
                self.sock.close()
            except Exception as ex:
                self.logger.warning(f"Close error: {ex}")
            self.sock = None

    def write_line(self, cmd):
        if not self.sock:
            raise IOError("Socket not open.")
        data = (cmd + "\n").encode("ascii")
        self.logger.debug(f"Sending: {cmd}")
        try:
            self.sock.sendall(data)
        except Exception as ex:
            self.logger.error(f"Error sending data: {ex}")
            raise

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
            except Exception as ex:
                self.logger.error(f"Recv error: {ex}")
                raise
            if not chunk:
                self.logger.warning("Socket recv empty data (closed?).")
                break

            line_buf.extend(chunk)
            if b'\n' in chunk:
                break

        line, _, _ = line_buf.partition(b'\n')
        msg = line.decode('ascii', errors='replace').strip()
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
    logger.info("=== APSM2000 LAN Mega Auto-Detect Script ===")
    logger.info(f"IP: {ip}:{port}")
    logger.info(f"CSV: {csv_file}")
    logger.info(f"Poll Interval: {poll_interval}s")
    logger.info(f"Debug: {debug}")

    # We'll try 9 variants total (CH1, A1, VPA1) x (none, :AC, :ACDC)
    #  1) "READ? VOLTS:CH1:AC,   VOLTS:CH2:AC,   VOLTS:CH3:AC"
    #  2) "READ? VOLTS:CH1,      VOLTS:CH2,      VOLTS:CH3"
    #  3) "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC"
    #  4) "READ? VOLTS:A1:AC,    VOLTS:A2:AC,    VOLTS:A3:AC"
    #  5) "READ? VOLTS:A1,       VOLTS:A2,       VOLTS:A3"
    #  6) "READ? VOLTS:A1:ACDC,  VOLTS:A2:ACDC,  VOLTS:A3:ACDC"
    #  7) "READ? VOLTS:VPA1:AC,  VOLTS:VPA2:AC,  VOLTS:VPA3:AC"
    #  8) "READ? VOLTS:VPA1,     VOLTS:VPA2,     VOLTS:VPA3"
    #  9) "READ? VOLTS:VPA1:ACDC,VOLTS:VPA2:ACDC,VOLTS:VPA3:ACDC"
    read_variants = [
        "READ? VOLTS:CH1:AC,   VOLTS:CH2:AC,   VOLTS:CH3:AC",
        "READ? VOLTS:CH1,      VOLTS:CH2,      VOLTS:CH3",
        "READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC",

        "READ? VOLTS:A1:AC,    VOLTS:A2:AC,    VOLTS:A3:AC",
        "READ? VOLTS:A1,       VOLTS:A2,       VOLTS:A3",
        "READ? VOLTS:A1:ACDC,  VOLTS:A2:ACDC,  VOLTS:A3:ACDC",

        "READ? VOLTS:VPA1:AC,  VOLTS:VPA2:AC,  VOLTS:VPA3:AC",
        "READ? VOLTS:VPA1,     VOLTS:VPA2,     VOLTS:VPA3",
        "READ? VOLTS:VPA1:ACDC,VOLTS:VPA2:ACDC,VOLTS:VPA3:ACDC",
    ]
    working_cmd = None

    m2000 = APSM2000_LAN(ip=ip, port=port, debug=debug)

    try:
        # 1) Connect
        m2000.open()

        # 2) *CLS
        m2000.write_line("*CLS")
        time.sleep(0.2)

        # 3) *IDN?
        m2000.write_line("*IDN?")
        time.sleep(0.1)
        try:
            idn = m2000.read_line(timeout_s=2.0)
            logger.info(f"IDN? => {idn}")
        except TimeoutError:
            logger.warning("No IDN response")

        # 4) Force multi-VPA mode
        m2000.write_line("MODE,1")
        time.sleep(0.2)
        m2000.write_line("SAVECONFIG")
        time.sleep(0.5)

        # 5) LOCKOUT => remote locked
        m2000.write_line("LOCKOUT")
        time.sleep(0.2)

        # 6) Try all 9 read variants
        for variant in read_variants:
            # remove extra whitespace
            cmd_clean = " ".join(variant.split())
            logger.info(f"Trying variant: {cmd_clean}")
            try:
                m2000.write_line(cmd_clean)
                resp_line = m2000.read_line(timeout_s=2.0)
                parts = resp_line.split(',')
                if len(parts) == 3:
                    # parse floats
                    testvals = [float(p) for p in parts]
                    logger.info(f"SUCCESS => '{cmd_clean}' => {testvals}")
                    working_cmd = cmd_clean
                    break
                else:
                    logger.warning(f"Wrong #fields => {resp_line}")
            except (TimeoutError, ValueError) as e:
                logger.warning(f"Variant '{cmd_clean}' failed: {e}")
            except Exception as ex:
                logger.error(f"Unexpected error with '{cmd_clean}': {ex}", exc_info=True)

        if not working_cmd:
            logger.error("No read command worked. Exiting.")
            return

        # 7) Stream loop => CSV
        with open(csv_file, "w", newline="") as cf:
            writer = csv.writer(cf)
            writer.writerow(["Timestamp", "Elapsed_s", "Channel1_V", "Channel2_V", "Channel3_V"])

            logger.info("Press Ctrl+C to stop streaming.\n")
            print(f"\n{'Time(s)':>10} | {'CH1':>10} | {'CH2':>10} | {'CH3':>10}")
            print("-" * 46)

            start_time = time.time()

            while True:
                try:
                    m2000.write_line(working_cmd)
                    line = m2000.read_line(timeout_s=2.0)
                    fields = line.split(',')
                    if len(fields) < 3:
                        logger.error(f"Incomplete => {line}")
                        continue
                    try:
                        v1 = float(fields[0])
                        v2 = float(fields[1])
                        v3 = float(fields[2])
                    except ValueError:
                        logger.error(f"Could not parse floats from '{line}'")
                        continue

                    elapsed = time.time() - start_time
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    print(f"{elapsed:10.2f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")
                    writer.writerow([now_str, f"{elapsed:.2f}", f"{v1:.3f}", f"{v2:.3f}", f"{v3:.3f}"])
                    cf.flush()

                    time.sleep(poll_interval)

                except TimeoutError:
                    logger.error("Timeout reading data.")
                except KeyboardInterrupt:
                    logger.info("Ctrl+C => stop streaming.")
                    break
                except Exception as loop_ex:
                    logger.error(f"Loop error: {loop_ex}", exc_info=True)

    except KeyboardInterrupt:
        logger.info("Ctrl+C during init => abort.")
    except Exception as main_ex:
        logger.error(f"Critical error: {main_ex}", exc_info=True)
    finally:
        # revert to local
        try:
            m2000.write_line("LOCAL")
            time.sleep(0.1)
        except:
            pass
        m2000.close()
        logger.info("=== Script done ===")


if __name__ == "__main__":
    IP_ADDR = "192.168.15.100"
    PORT    = 10733
    CSV_OUT = "apsm2000_lan_datalog.csv"
    SEC     = 1.0
    DBG     = False

    run_lan_script(
        ip=IP_ADDR,
        port=PORT,
        csv_file=CSV_OUT,
        poll_interval=SEC,
        debug=DBG
    )

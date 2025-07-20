#!/usr/bin/env python3
"""
APSM2000 Basic Communication - Reads 3 Channels
Uses either ACDC or no subfield to avoid timeouts if :AC is not recognized.
"""

import socket
import time
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

def main():
    HOST = "192.168.15.100"
    PORT = 10733

    logger.info(f"Attempting connection to {HOST}:{PORT}")

    s = None
    try:
        # Create socket & connect
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(10.0)
        s.connect((HOST, PORT))
        logger.info("Connected successfully")
        time.sleep(1.0)

        # 1) Clear
        logger.info("Sending *CLS")
        s.sendall(b"*CLS\n")
        time.sleep(0.5)

        # 2) IDN?
        logger.info("Sending *IDN?")
        s.sendall(b"*IDN?\n")
        time.sleep(0.5)
        try:
            idn_resp = s.recv(4096)
            if idn_resp:
                logger.info("IDN Reply: " + idn_resp.decode('ascii', errors='ignore').strip())
            else:
                logger.error("No IDN response.")
                return
        except socket.timeout:
            logger.error("Timeout waiting for IDN reply.")
            return

        # 3) Read channels one at a time
        channels = [1, 2, 3]
        for ch in channels:
            command_str = f"READ? VOLTS:CH{ch}:AC"
            logger.info(f"Sending: {command_str}")
            s.sendall((command_str + "\n").encode('ascii'))
            time.sleep(0.5)  # small delay between commands
            
            try:
                resp = s.recv(4096)
                if resp:
                    resp_str = resp.decode('ascii', errors='ignore').strip()
                    logger.info(f"Raw response for CH{ch}: {resp_str}")
                    try:
                        voltage_val = float(resp_str)
                        logger.info(f"Channel {ch} AC Voltage: {voltage_val:.6f} V")
                    except ValueError:
                        logger.error(f"Could not parse voltage for CH{ch}: '{resp_str}'")
                else:
                    logger.error(f"No response for CH{ch} read command")
            except socket.timeout:
                logger.error(f"Timeout reading channel {ch} voltage")

    except socket.timeout:
        logger.error("Connection timed out.")
    except ConnectionRefusedError:
        logger.error("Connection refused - check IP/Port or M2000 settings.")
    except Exception as ex:
        logger.error(f"Unexpected error: {ex}")
    finally:
        if s:
            try:
                # Return to local
                logger.info("Sending LOCAL")
                s.sendall(b"LOCAL\n")
                time.sleep(0.3)
                s.close()
                logger.info("Connection closed")
            except:
                pass

if __name__ == "__main__":
    main()

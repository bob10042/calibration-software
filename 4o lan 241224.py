#!/usr/bin/env python3
"""
APSM2000 Basic Communication Test - Debugged
Reads 3 channels of AC voltage (CH1, CH2, CH3) using M2000 commands.
"""

import socket
import time
import sys
import logging

# Setup logging (DEBUG level for detail; change to INFO if you prefer less verbosity)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

def test_communication():
    # Connection settings
    HOST = "192.168.15.100"  # <-- Adjust to your M2000's IP
    PORT = 10733            # M2000 default LAN port
    
    logger.info(f"Attempting connection to {HOST}:{PORT}")
    
    s = None
    try:
        # 1) Create and connect the socket
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(10.0)  # 10 second timeout
        s.connect((HOST, PORT))
        logger.info("Connected successfully")
        time.sleep(1.0)

        # 2) Clear status
        logger.info("Sending *CLS")
        s.sendall(b"*CLS\n")
        time.sleep(0.5)

        # 3) Query IDN
        logger.info("Sending *IDN?")
        s.sendall(b"*IDN?\n")
        time.sleep(0.5)
        try:
            response = s.recv(4096)
            if response:
                idn_reply = response.decode('ascii', errors='ignore').strip()
                logger.info(f"IDN Reply: {idn_reply}")
            else:
                logger.error("No response received from IDN query")
                return
        except socket.timeout:
            logger.error("Timeout while waiting for IDN response")
            return

        # 4) Read AC voltage from channels 1..3
        #    Command syntax: READ? VOLTS:CHx:AC
        #    This returns one floating-point value (in ASCII)
        #    If you have AC+DC or want DC only, use :ACDC or :DC instead
        channels = [1, 2, 3]  # only 3 channels
        for ch in channels:
            command_str = f"READ? VOLTS:CH{ch}:AC"
            logger.info(f"Sending: {command_str}")
            s.sendall((command_str + "\n").encode('ascii'))
            
            time.sleep(0.5)  # small delay to wait for M2000 to process
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
        logger.error("Connection timed out")
    except ConnectionRefusedError:
        logger.error("Connection refused - verify IP address and port")
    except Exception as e:
        logger.error(f"Error: {str(e)}")
    finally:
        # 5) Attempt to return to LOCAL and close
        if s:
            try:
                logger.info("Sending LOCAL")
                s.sendall(b"LOCAL\n")
                time.sleep(0.5)
                s.close()
                logger.info("Connection closed")
            except Exception:
                pass

if __name__ == "__main__":
    test_communication()

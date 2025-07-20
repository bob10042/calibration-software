#!/usr/bin/env python3
"""
APSM2000 Basic Communication Test
WORKING INITIALIZATION VERSION - 24/12/2024
Successfully establishes connection and gets device identification
"""

import socket
import time
import sys
import logging

# Setup logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

def test_communication():
    # Connection settings
    HOST = "192.168.15.100"
    PORT = 10733
    
    logger.info(f"Attempting connection to {HOST}:{PORT}")
    
    try:
        # Create socket
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(10.0)  # 10 second timeout
        
        # Connect
        s.connect((HOST, PORT))
        logger.info("Connected successfully")
        time.sleep(1.0)  # Wait after connection
        
        # Clear status
        logger.info("Sending *CLS")
        s.sendall(b"*CLS\n")
        time.sleep(0.5)
        
        # Get identification
        logger.info("Sending *IDN?")
        s.sendall(b"*IDN?\n")
        time.sleep(1.0)  # Longer delay before receiving response
        
        try:
            response = s.recv(4096)
            if response:
                logger.info(f"IDN Reply: {response.decode('ascii', errors='ignore').strip()}")
            else:
                logger.error("No response received from IDN query")
        except socket.timeout:
            logger.error("Timeout while waiting for IDN response")
            return

        # Read voltage measurements for each channel
        channels = [1, 2, 3, 4]
        for channel in channels:
            try:
                # Select channel
                logger.info(f"Selecting channel {channel}")
                s.sendall(f"CHAN {channel}\n".encode())
                time.sleep(0.5)
                
                # Read voltage
                logger.info(f"Reading voltage for channel {channel}")
                s.sendall(b"MEAS:VOLT?\n")
                time.sleep(0.5)
                response = s.recv(4096)
                if response:
                    voltage = float(response.decode('ascii', errors='ignore').strip())
                    logger.info(f"Channel {channel} Voltage: {voltage:.6f} V")
                else:
                    logger.error(f"No response received for channel {channel} voltage query")
            except (socket.timeout, ValueError) as e:
                logger.error(f"Error reading channel {channel} voltage: {str(e)}")
        
    except socket.timeout:
        logger.error("Connection timed out")
    except ConnectionRefusedError:
        logger.error("Connection refused - verify IP address and port")
    except Exception as e:
        logger.error(f"Error: {str(e)}")
    finally:
        try:
            s.sendall(b"LOCAL\n")
            s.close()
            logger.info("Connection closed")
        except:
            pass

if __name__ == "__main__":
    test_communication()
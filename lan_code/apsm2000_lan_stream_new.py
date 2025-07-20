#!/usr/bin/env python3
"""
APSM2000 Basic Communication Test
Testing simpler voltage read commands
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
        s.settimeout(5.0)  # 5 second timeout
        
        # Connect
        s.connect((HOST, PORT))
        logger.info("Connected successfully")
        
        # Clear status
        logger.info("Sending *CLS")
        s.sendall(b"*CLS\n")
        time.sleep(0.1)
        
        # Get identification
        logger.info("Sending *IDN?")
        s.sendall(b"*IDN?\n")
        response = s.recv(4096)
        logger.info(f"IDN Reply: {response.decode('ascii', errors='ignore').strip()}")
        time.sleep(0.1)
        
        # Try simpler voltage read commands
        
        # Channel 1
        logger.info("Selecting channel 1")
        s.sendall(b"CHAN 1\n")
        time.sleep(0.1)
        
        logger.info("Reading voltage from channel 1")
        s.sendall(b"VOLT?\n")  # Simplified command
        response = s.recv(4096)
        logger.info(f"CH1 Voltage Reply: {response.decode('ascii', errors='ignore').strip()}")
        time.sleep(0.1)
        
        # Try alternative format if first one fails
        if not response:
            logger.info("Trying alternative voltage read format")
            s.sendall(b"V?\n")  # Even simpler command
            response = s.recv(4096)
            logger.info(f"CH1 Alternative Voltage Reply: {response.decode('ascii', errors='ignore').strip()}")
            time.sleep(0.1)
        
        # If we get a response, try reading all channels at once
        logger.info("Trying multi-channel read")
        s.sendall(b"READ? V1,V2,V3\n")  # Alternative multi-channel format
        response = s.recv(4096)
        logger.info(f"Multi-channel Reply: {response.decode('ascii', errors='ignore').strip()}")
        
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
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
import csv
from datetime import datetime

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

def send_command(sock, command, expect_reply=False):
    """Helper function to send commands and receive replies"""
    try:
        logger.info(f"Sending: {command}")
        command_bytes = f"{command}\n".encode('ascii')
        sock.sendall(command_bytes)
        
        # Delay for instrument processing
        time.sleep(0.5)
        
        if expect_reply:
            try:
                response = sock.recv(4096)
                reply = response.decode('ascii', errors='ignore').strip()
                logger.info(f"Received: {reply}")
                return reply
            except socket.timeout:
                logger.error(f"Timeout while waiting for reply to: {command}")
                return None
            except Exception as e:
                logger.error(f"Error reading reply to {command}: {str(e)}")
                return None
        return None
    except socket.timeout:
        logger.error(f"Timeout while sending command: {command}")
        return None
    except Exception as e:
        logger.error(f"Error sending command {command}: {str(e)}")
        return None

def test_communication():
    # Connection settings
    HOST = "192.168.15.100"
    PORT = 10733
    
    logger.info(f"Attempting connection to {HOST}:{PORT}")
    s = None
    
    try:
        # Create socket
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(5.0)  # 5 second timeout
        
        # Connect
        s.connect((HOST, PORT))
        logger.info("Connected successfully")
        time.sleep(2.0)  # Longer delay after connection
        
        # Get identification first (this worked before)
        idn = send_command(s, "*IDN?", expect_reply=True)
        if not idn:
            logger.error("Failed to get device identification")
            return
        time.sleep(1.0)
        
        # Only proceed if IDN worked
        logger.info("Device identified successfully, attempting voltage reading")
        
        # Open CSV for logging
        csv_filename = f"voltage_readings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        try:
            with open(csv_filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Timestamp', 'CH1 V', 'CH2 V', 'CH3 V'])
                
                logger.info("\nStarting voltage readings (Press Ctrl+C to stop)...")
                print(f"\n{'Time':^12} | {'CH1 V':^10} | {'CH2 V':^10} | {'CH3 V':^10}")
                print("-" * 50)
                
                start_time = time.time()
                
                try:
                    while True:
                        # Read each channel individually
                        v1 = send_command(s, "READ? VOLTS:CH1", expect_reply=True)
                        time.sleep(0.5)
                        v2 = send_command(s, "READ? VOLTS:CH2", expect_reply=True)
                        time.sleep(0.5)
                        v3 = send_command(s, "READ? VOLTS:CH3", expect_reply=True)
                        
                        if v1 and v2 and v3:  # Only process if we got all readings
                            try:
                                # Parse values
                                volts1 = float(v1)
                                volts2 = float(v2)
                                volts3 = float(v3)
                                
                                timestamp = datetime.now()
                                elapsed = time.time() - start_time
                                
                                # Print to console
                                print(f"{elapsed:>8.1f}s | {volts1:>10.3f} | {volts2:>10.3f} | {volts3:>10.3f}")
                                
                                # Write to CSV
                                writer.writerow([timestamp.strftime('%Y-%m-%d %H:%M:%S'), 
                                              volts1, volts2, volts3])
                                csvfile.flush()
                                
                            except ValueError as e:
                                logger.error(f"Error parsing voltage values: {e}")
                        else:
                            logger.warning("Failed to read one or more channels")
                        
                        time.sleep(1.0)  # Wait before next readings
                        
                except KeyboardInterrupt:
                    logger.info("\nUser stopped readings")
                except Exception as e:
                    logger.error(f"Error in reading loop: {str(e)}")
                
        except IOError as e:
            logger.error(f"Error with CSV file {csv_filename}: {str(e)}")
        except Exception as e:
            logger.error(f"Unexpected error with CSV handling: {str(e)}")
            
    except socket.timeout:
        logger.error("Connection timed out")
    except ConnectionRefusedError:
        logger.error("Connection refused - verify IP address and port")
    except Exception as e:
        logger.error(f"Error: {str(e)}")
    finally:
        if s:
            try:
                send_command(s, "LOCAL")
                s.close()
                logger.info("Connection closed")
            except Exception as e:
                logger.error(f"Error closing connection: {str(e)}")

if __name__ == "__main__":
    test_communication()
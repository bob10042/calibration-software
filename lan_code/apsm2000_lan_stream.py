#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) LAN/TCP Streaming Example with Built-In Logging & Error Handling

Features:
- Opens a TCP socket connection to the APSM2000 (default port 10733).
- Continuously queries AC/DC voltages from 3 channels.
- Prints readings to console (3 decimal places) and logs to CSV file.
- Incorporates Python's `logging` module for debug/error tracking.
- Graceful exception handling.

Usage:
  python apsm2000_lan_stream.py
    (Creates "apsm2000_lan_datalog.csv" with results and prints logs to console.)
"""

import socket
import time
import csv
import sys
import logging
import os
from datetime import datetime

####################
# 1) Logging Setup
####################
def setup_logger(debug=False):
    """
    Sets up a global logger with console output.
    If debug=True, sets level to DEBUG; otherwise INFO.
    """
    logger = logging.getLogger("APSM2000_LAN_Logger")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    # Clear old handlers
    logger.handlers.clear()
    
    # Console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    
    # Format
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    
    return logger

# Create module-level logger
logger = setup_logger(debug=True)  # Default to debug for troubleshooting

#######################################
# 2) APSM2000 LAN Communication Class
#######################################
class APSM2000_LAN:
    """
    Handles TCP/IP communication with the APSM2000.
    Provides connect/disconnect, send_command, read_response methods.
    """
    def __init__(self, host="192.168.15.100", port=10733):
        self.host = host
        self.port = port
        self.socket = None
        self.timeout = 2.0  # Reduced timeout to 2 seconds
        
    def connect(self):
        """Opens TCP connection to APSM2000 and initializes the device."""
        logger.info(f"Connecting to APSM2000 at {self.host}:{self.port}...")
        
        try:
            # Create and connect socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(self.timeout)
            self.socket.connect((self.host, self.port))
            logger.info("Connected successfully.")
            time.sleep(0.1)  # Brief delay after connection
            
            # Basic initialization sequence
            logger.debug("Initializing device...")
            
            # Clear any previous errors
            self.send_command("*CLS")
            time.sleep(0.1)
            
            # Reset device
            self.send_command("*RST")
            time.sleep(0.1)
            
            # Set remote mode
            self.send_command("SYSTEM:REMOTE")
            time.sleep(0.1)
            
            # Configure channels exactly like lan_tcpip.py
            logger.debug("Configuring channels...")
            for ch in range(1, 4):
                self.send_command(f"MEAS:VOLTAGE:ACDC CH{ch}")
                time.sleep(0.1)
            
            # Set integration period
            logger.debug("Setting integration period...")
            self.send_command("CONF:INTEGRATION 0.01")
            time.sleep(0.1)
            
        except socket.timeout:
            logger.error("Connection timed out")
            raise
        except ConnectionRefusedError:
            logger.error(f"Connection refused - check if APSM2000 is ready at {self.host}:{self.port}")
            raise
        except Exception as e:
            logger.error(f"Connection error: {str(e)}")
            raise
            
    def disconnect(self):
        """Closes TCP connection."""
        if self.socket:
            self.socket.close()
            self.socket = None
            logger.info("Disconnected from APSM2000.")
            
    def send_command(self, command, expect_response=False):
        """Helper function to send a command and get response if needed"""
        if not self.socket:
            raise IOError("Not connected")
            
        try:
            logger.debug(f"Sending: {command}")
            self.socket.sendall((command + "\n").encode())
            time.sleep(0.1)  # Same delay as working script
            
            if expect_response:
                try:
                    response = self.socket.recv(1024).decode().strip()
                    if response:
                        logger.debug(f"Response: {response}")
                        return response
                    if response.startswith("ERR"):
                        logger.warning(f"Device error: {response}")
                except Exception as e:
                    logger.error(f"Error: {e}")
            return None
            
        except Exception as e:
            logger.error(f"Command error: {str(e)}")
            raise

##############################################
# 3) Main Streaming and Data Logging Function
##############################################
def stream_voltages_and_log(
    host="192.168.15.100",
    port=10733,
    output_csv="apsm2000_lan_datalog.csv",
    poll_interval=1.0,
    debug=True  # Default to debug for troubleshooting
):
    """
    1) Connects to APSM2000 via LAN/TCP.
    2) Continuously queries AC/DC voltages from 3 channels.
    3) Prints to console (3 decimal places).
    4) Logs to CSV with timestamps.
    5) Handles errors gracefully.
    6) Stops on Ctrl+C.
    """
    # Set logging level
    global logger
    logger = setup_logger(debug=debug)
    
    logger.info("=== Starting APSM2000 LAN Streaming ===")
    logger.info(f"Target: {host}:{port}")
    logger.info(f"Output CSV: {output_csv}")
    logger.info(f"Poll Interval: {poll_interval}s")
    logger.info(f"Debug Mode: {debug}")
    
    # Create APSM2000 connection
    aps = APSM2000_LAN(host=host, port=port)
    
    try:
        # Connect and verify communication
        aps.connect()
        time.sleep(0.5)  # Settle time after connect
        
        # Optional: lock front panel
        # aps.send_command("LOCKOUT")
        
        # Prepare CSV file
        with open(output_csv, mode='w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            # Write header
            writer.writerow([
                "Timestamp",
                "Elapsed (s)", 
                "Voltage1 (AC+DC)",
                "Voltage2 (AC+DC)",
                "Voltage3 (AC+DC)"
            ])
            
            logger.info(f"CSV logging started: {output_csv}")
            start_time = time.time()
            
            # Print header for console output
            print(f"\n{'Time(s)':>8} | {'V1':>10} | {'V2':>10} | {'V3':>10}")
            print("-" * 46)
            
            while True:
                try:
                    # Read voltages using the exact working pattern from lan_tcpip.py
                    measurements = {}
                    
                    for ch in range(1, 4):
                        # Direct measurement query like in lan_tcpip.py
                        cmd = f"MEAS:VOLTAGE:ACDC? CH{ch}"
                        logger.debug(f"Reading CH{ch}: {cmd}")
                        response = aps.send_command(cmd, expect_response=True)
                        time.sleep(0.1)  # Delay after each measurement
                        if response:
                            try:
                                measurements[ch] = float(response)
                                logger.debug(f"CH{ch} voltage: {measurements[ch]:.3f}V")
                            except ValueError:
                                logger.warning(f"Parse error for CH{ch}: {response}")
                                break
                        else:
                            logger.warning(f"No response from CH{ch}")
                            break
                            
                        time.sleep(0.1)  # Consistent delay between measurements
                    
                    # Skip if we didn't get all measurements
                    if len(measurements) != 3:
                        logger.warning("Incomplete measurements, retrying...")
                        continue
                    
                    # Get timestamps
                    now = datetime.now()
                    elapsed = time.time() - start_time
                    
                    # Print to console (3 decimal places)
                    print(f"{elapsed:8.2f} | {measurements[1]:10.3f} | {measurements[2]:10.3f} | {measurements[3]:10.3f}")
                    
                    # Write to CSV
                    writer.writerow([
                        now.strftime("%Y-%m-%d %H:%M:%S"),
                        f"{elapsed:.2f}",
                        f"{measurements[1]:.3f}",
                        f"{measurements[2]:.3f}",
                        f"{measurements[3]:.3f}"
                    ])
                    csvfile.flush()  # Ensure data is written
                    
                    # Wait for next poll
                    time.sleep(poll_interval)
                    
                except (socket.timeout, TimeoutError) as te:
                    logger.error(f"Timeout during streaming: {str(te)}")
                    # Optional: try to recover connection here
                    time.sleep(1.0)  # Wait before retry
                except Exception as e:
                    logger.error(f"Error during streaming: {str(e)}")
                    time.sleep(1.0)  # Wait before retry
                    
    except KeyboardInterrupt:
        logger.info("\nUser stopped streaming (Ctrl+C)")
    except Exception as e:
        logger.error(f"Critical error: {str(e)}", exc_info=True)
    finally:
        # Optionally restore local control
        try:
            if aps.socket:
                aps.send_command("LOCAL")
        except:
            pass
        
        # Close connection
        aps.disconnect()
        logger.info("=== APSM2000 LAN streaming stopped ===")

if __name__ == "__main__":
    # Default settings - adjust as needed
    HOST = "192.168.15.100"  # Your APSM2000's IP
    PORT = 10733             # Default port
    OUTPUT_CSV = "apsm2000_lan_datalog.csv"
    POLL_INTERVAL = 1.0      # seconds
    DEBUG_MODE = True        # Enable debug logging for troubleshooting
    
    stream_voltages_and_log(
        host=HOST,
        port=PORT,
        output_csv=OUTPUT_CSV,
        poll_interval=POLL_INTERVAL,
        debug=DEBUG_MODE
    )
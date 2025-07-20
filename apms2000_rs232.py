#!/usr/bin/env python3
"""
APSM2000 (M2000 Series) RS232 Communication Example
(Also works with USB-to-RS232 adapters)

Features:
- Opens RS232 port (or USB-to-RS232 adapter's virtual COM port).
- Configures for proper baud rate, flow control, etc.
- Streams voltage measurements from 3 channels.
- Logs data to CSV file.
- Includes basic error handling.

Requirements:
- Python 3.x
- pyserial (pip install pyserial)
- RS232 null-modem cable or USB-to-RS232 adapter
- APSM2000 configured for matching baud rate and handshake
"""

import serial
import time
import csv
import sys
import logging
from datetime import datetime


def setup_logger(debug=False):
    """Configure logging to console."""
    logger = logging.getLogger("APSM2000_RS232")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    
    # Clear any existing handlers
    logger.handlers.clear()
    
    # Console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    
    # Format
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    
    return logger


class APSM2000_RS232:
    """
    Handles RS232 communication with APSM2000.
    Works with direct RS232 or USB-to-RS232 adapters.
    """
    
    def __init__(self, port="COM1", baud_rate=115200, timeout=1.0):
        self.logger = setup_logger()
        self.port = port
        self.baud_rate = baud_rate
        self.timeout = timeout
        self.ser = None
    
    def open(self):
        """Opens the serial port with proper settings."""
        self.logger.info(f"Opening {self.port} at {self.baud_rate} baud...")
        
        try:
            self.ser = serial.Serial()
            self.ser.port = self.port
            self.ser.baudrate = self.baud_rate
            
            # 8 data bits, no parity, 1 stop bit
            self.ser.bytesize = serial.EIGHTBITS
            self.ser.parity = serial.PARITY_NONE
            self.ser.stopbits = serial.STOPBITS_ONE
            
            # Read timeout
            self.ser.timeout = self.timeout
            
            # Hardware flow control (if APSM2000 is configured for it)
            self.ser.rtscts = True
            
            # DTR must be asserted for M2000 to see controller
            self.ser.dtr = True
            
            self.ser.open()
            self.logger.info(f"Opened {self.port} successfully.")
            
        except serial.SerialException as e:
            self.logger.error(f"Failed to open {self.port}: {str(e)}")
            raise
    
    def close(self):
        """Closes the serial port."""
        if self.ser and self.ser.is_open:
            self.logger.info(f"Closing {self.port}...")
            self.ser.close()
            self.logger.info(f"Closed {self.port}.")
    
    def write_line(self, command):
        """Sends ASCII command with newline."""
        if not self.ser or not self.ser.is_open:
            raise IOError("Port not open")
        
        try:
            # Add newline if needed
            if not command.endswith('\n'):
                command += '\n'
            
            # Convert to bytes and send
            self.ser.write(command.encode('ascii'))
            self.logger.debug(f"Sent: {command.strip()}")
            
        except serial.SerialException as e:
            self.logger.error(f"Write error: {str(e)}")
            raise
    
    def read_line(self):
        """
        Reads until newline or timeout.
        Returns stripped string or None if timeout/error.
        """
        if not self.ser or not self.ser.is_open:
            raise IOError("Port not open")
        
        try:
            # readline() already handles the timeout we set
            line = self.ser.readline()
            
            if not line:  # timeout
                self.logger.warning("Read timeout")
                return None
            
            # Decode and strip whitespace/newlines
            decoded = line.decode('ascii').strip()
            self.logger.debug(f"Received: {decoded}")
            return decoded
            
        except serial.SerialException as e:
            self.logger.error(f"Read error: {str(e)}")
            raise


def stream_voltages(
    port="COM1",
    baud_rate=115200,
    output_csv="apsm2000_rs232_log.csv",
    poll_interval=1.0,
    debug=False
):
    """
    Opens RS232 connection to APSM2000 and streams voltage measurements.
    
    Args:
        port: Serial port name (e.g., "COM1", "/dev/ttyUSB0")
        baud_rate: Must match APSM2000's setting
        output_csv: Where to save the measurements
        poll_interval: Seconds between readings
        debug: True for verbose logging
    """
    
    # Set up logging
    logger = setup_logger(debug=debug)
    logger.info("=== Starting APSM2000 RS232 Streaming ===")
    logger.info(f"Port: {port}")
    logger.info(f"Baud Rate: {baud_rate}")
    logger.info(f"Output CSV: {output_csv}")
    logger.info(f"Poll Interval: {poll_interval}s")
    
    # Create device object
    m2000 = APSM2000_RS232(
        port=port,
        baud_rate=baud_rate,
        timeout=2.0
    )
    
    try:
        # Open port
        m2000.open()
        
        # Clear interface
        m2000.write_line("*CLS")
        time.sleep(0.2)
        
        # Optional: get ID
        m2000.write_line("*IDN?")
        idn = m2000.read_line()
        if idn:
            logger.info(f"Device ID: {idn}")
        
        # Optional: lock front panel
        # m2000.write_line("LOCKOUT")
        
        # Prepare CSV file
        with open(output_csv, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([
                "Timestamp",
                "Elapsed (s)",
                "CH1 V(AC+DC)",
                "CH2 V(AC+DC)",
                "CH3 V(AC+DC)"
            ])
            
            # Print header
            print("\nStreaming voltages (Ctrl+C to stop)...")
            print(f"{'Time(s)':>8} | {'CH1(V)':>10} | {'CH2(V)':>10} | {'CH3(V)':>10}")
            print("-" * 46)
            
            # Start time for elapsed calculation
            start_time = time.time()
            
            try:
                while True:
                    # Request 3-channel voltage reading
                    m2000.write_line("READ? VOLTS:CH1:ACDC, VOLTS:CH2:ACDC, VOLTS:CH3:ACDC")
                    response = m2000.read_line()
                    
                    if not response:
                        logger.warning("No response to voltage query")
                        continue
                    
                    # Parse the comma-separated values
                    try:
                        v1, v2, v3 = map(float, response.split(','))
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Could not parse response: {response} => {e}")
                        continue
                    
                    # Get timestamps
                    now = datetime.now()
                    elapsed = time.time() - start_time
                    
                    # Print to console (3 decimal places)
                    print(f"{elapsed:8.1f} | {v1:10.3f} | {v2:10.3f} | {v3:10.3f}")
                    
                    # Write to CSV
                    writer.writerow([
                        now.strftime("%Y-%m-%d %H:%M:%S"),
                        f"{elapsed:.1f}",
                        f"{v1:.3f}",
                        f"{v2:.3f}",
                        f"{v3:.3f}"
                    ])
                    csvfile.flush()  # Ensure it's written
                    
                    # Wait for next poll
                    time.sleep(poll_interval)
                    
            except KeyboardInterrupt:
                print("\nUser stopped streaming.")
                
    except Exception as e:
        logger.error(f"Error during streaming: {str(e)}")
        
    finally:
        # Optionally restore local control
        # m2000.write_line("LOCAL")
        
        # Always close the port
        m2000.close()
        logger.info("=== Streaming stopped ===")


if __name__ == "__main__":
    # Adjust these settings as needed
    PORT = "COM1"          # or "COM4", "/dev/ttyUSB0", etc.
    BAUD_RATE = 115200    # must match APSM2000's setting
    OUTPUT_CSV = "apsm2000_rs232_log.csv"
    POLL_INTERVAL = 1.0   # seconds between readings
    DEBUG = False         # set True for verbose logging
    
    # Start streaming
    stream_voltages(
        port=PORT,
        baud_rate=BAUD_RATE,
        output_csv=OUTPUT_CSV,
        poll_interval=POLL_INTERVAL,
        debug=DEBUG
    )
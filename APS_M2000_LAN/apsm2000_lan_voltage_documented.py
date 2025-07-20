"""
APS M2000 Power Analyzer LAN Communication Module

This module provides a class for communicating with the APS M2000 Power Analyzer over LAN/Ethernet.
It handles device connection, command sending, and voltage measurements across multiple channels.

Protocol Details:
- Uses TCP/IP for communication
- Commands must end with \r\n (CRLF)
- Default port is 10733
- Requires 500ms delay between commands
- 1 second delay between measurement cycles
"""

import socket
import time
import sys

class APSM2000:
    """
    A class to handle communication with the APS M2000 Power Analyzer over LAN.
    
    This class manages the TCP/IP connection to the device and provides methods
    for sending commands and reading measurements.
    
    Attributes:
        ip (str): IP address of the APS M2000 device
        port (int): TCP port number for communication (default: 10733)
        socket: TCP socket for device communication
    """

    def __init__(self, ip, port=10733):
        """
        Initialize connection parameters for APS M2000.

        Args:
            ip (str): IP address of the device
            port (int, optional): TCP port number. Defaults to 10733.

        Socket Configuration:
            - Uses TCP (SOCK_STREAM)
            - 3 second timeout for commands
            - Automatically handles connection cleanup
        """
        self.ip = ip
        self.port = port
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.settimeout(3.0)  # 3 second timeout for commands

    def connect(self):
        """
        Establish connection to the device.

        This method:
        1. Cleans up any existing connection
        2. Creates a new socket
        3. Connects to the device
        4. Clears any pending data
        5. Waits for port stabilization

        Returns:
            bool: True if connection successful, False otherwise

        Note:
            Includes a 2-second stabilization delay after connection
        """
        try:
            # Clear any existing connection
            try:
                self.socket.close()
            except:
                pass
            
            # Create new socket and connect
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(3.0)  # 3 second timeout
            self.socket.connect((self.ip, self.port))
            
            # Clear any pending data
            try:
                while self.socket.recv(1024):
                    pass
            except socket.error:
                pass
            
            # Give time for port to stabilize
            time.sleep(2)  # 2 second stabilization time
                
            return True
        except socket.error as e:
            print(f"Connection error: {e}")
            return False

    def disconnect(self):
        """
        Close the connection to the device.

        This method ensures proper cleanup of the socket connection.
        Any errors during closing are logged but not raised.
        """
        try:
            self.socket.close()
        except socket.error as e:
            print(f"Error closing connection: {e}")

    def send_command(self, cmd):
        """
        Send a command to the device and return the response.

        This method handles:
        1. Buffer clearing before sending
        2. Command termination (\r\n)
        3. Command sending with debug output
        4. Response reading with timeout
        5. Response processing

        Args:
            cmd (str): Command to send to the device

        Returns:
            str: Device response stripped of whitespace, or None if error

        Note:
            - Includes 200ms delay after sending command
            - 3 second timeout for response
            - Prints debug information for commands and responses
        """
        try:
            # Clear any pending data before sending command
            try:
                while self.socket.recv(1024):
                    pass
            except socket.error:
                pass

            # Add termination if not present
            if not cmd.endswith('\r\n'):
                cmd += '\r\n'
            
            print(f"Sending: {cmd.strip()}")  # Debug print
            
            # Send command
            self.socket.sendall(cmd.encode())
            time.sleep(0.2)  # 200ms delay after sending command
            
            # Read response with timeout
            response = ''
            start_time = time.time()
            while True:
                try:
                    data = self.socket.recv(1024).decode()
                    if data:
                        response += data
                        if '\n' in data:  # Got complete response
                            break
                    elif response:  # No more data but have partial response
                        break
                    # Timeout after 3 seconds
                    if time.time() - start_time > 3:
                        print(f"Timeout waiting for response to: {cmd.strip()}")
                        break
                except socket.error:
                    break
            
            response = response.strip()
            print(f"Received: {response}")  # Debug print
            return response
        except socket.error as e:
            print(f"Error sending command '{cmd.strip()}': {e}")
            return None

    def check_error(self):
        """
        Check for any device errors.

        Sends the *ERR? command to query device error status.

        Returns:
            str: Error code from device, "0" indicates no error
        """
        return self.send_command('*ERR?')

    def format_reading(self, value, measurement_type):
        """
        Format measurement values into human-readable format.

        This method handles:
        1. Scientific notation parsing
        2. Unit conversion (e.g., V to mV)
        3. Proper decimal formatting

        Args:
            value (str): Raw measurement value from device
            measurement_type (str): Type of measurement (e.g., "voltage")

        Returns:
            str: Formatted measurement with appropriate units

        Example:
            "1.23E-3" -> "1.230 mV" for voltage measurements
        """
        try:
            if not value:
                return "No reading"
            
            # Split the value into its components (handles multiple E notation values)
            parts = value.strip()
            
            # Find all numbers with scientific notation
            values = []
            current_num = ""
            for char in parts:
                if char in '+-' and current_num:
                    if 'E' in current_num:  # Only split if we've completed an E notation number
                        values.append(current_num)
                        current_num = char
                    else:
                        current_num += char
                else:
                    current_num += char
            values.append(current_num)  # Add the last number
            
            # Convert each value from scientific notation
            converted_values = []
            for val in values:
                if 'E' in val:
                    mantissa, exponent = val.split('E')
                    converted_values.append(float(mantissa) * (10 ** int(exponent)))
                else:
                    converted_values.append(float(val))
            
            # Use the first value as the main measurement
            main_value = converted_values[0]
            
            # Format based on measurement type
            if measurement_type == "voltage":
                if abs(main_value) < 1:
                    return f"{main_value * 1000:.1f} mV"
                return f"{main_value:.3f} V"
            else:
                return f"{main_value:.3f}"
                
        except Exception as e:
            print(f"Error formatting {value}: {e}")
            return value

    def read_voltages(self):
        """
        Read voltage measurements from all channels.

        This method:
        1. Iterates through channels 1-3
        2. Checks channel availability
        3. Takes multiple readings over a 15-second period
        4. Formats and displays measurements

        Channel Reading Process:
        1. Clear input buffers
        2. Query channel status
        3. Check for errors
        4. If channel available, take readings
        5. Wait between readings

        Note:
            - 3 second delay between readings
            - Readings continue for 15 seconds per channel
            - Includes error checking after each channel operation
        """
        try:
            print("\n" + "="*50)
            print(f"Reading started at: {time.strftime('%H:%M:%S')}")
            print("="*50)
            
            # Read channels 1-3
            for channel in range(1, 4):
                # Clear any pending data before reading channel
                try:
                    while self.socket.recv(1024):
                        pass
                except socket.error:
                    pass
                
                # Check if channel is available
                channel_info = self.send_command(f'CHNL?,{channel}')
                error = self.check_error()  # Check for errors after channel selection
                
                if error and error != "0":
                    print(f"Error on channel {channel}: {error}")
                    continue
                    
                if channel_info and channel_info != "NF0":  # Skip if channel not found
                    print(f"\nReading Channel {channel} (Status: {channel_info})")
                    print(f"Hold time: 15 seconds")
                    print("-"*30)
                    
                    # Take multiple readings over the hold period
                    start_time = time.time()
                    readings_count = 0
                    
                    while time.time() - start_time < 15:  # 15 second hold time
                        readings_count += 1
                        print(f"\nReading #{readings_count}")
                        
                        # Read voltage with proper format
                        volts = self.send_command(f'READ?,VOLTS,CH{channel},ACDC')  # RMS voltage
                        print(f"Voltage: {self.format_reading(volts, 'voltage')}")
                        
                        time.sleep(3)  # Wait 3 seconds between readings
                else:
                    print(f"\nChannel {channel}: Not available or not found")
                
        except Exception as e:
            print(f"Error reading voltages: {e}")

def main():
    """
    Main program entry point.

    This function:
    1. Sets up device connection
    2. Initializes communication
    3. Gets device information
    4. Reads configuration
    5. Performs voltage measurements
    6. Handles cleanup

    Device Configuration:
    - IP: 192.168.1.101
    - Port: 10733 (default)

    Note:
        Includes proper error handling and connection cleanup
    """
    # Device configuration
    ip = "192.168.1.101"  # IP address of the APS M2000
    port = 10733  # Default port for instrument communication
    
    try:
        # Initialize and connect to device
        print(f"Connecting to APS M2000 at {ip}:{port}...")
        device = APSM2000(ip, port)
        
        if not device.connect():
            print("Failed to connect to device")
            sys.exit(1)
            
        print("Connected successfully")
        
        # Get device identification
        print("\nDevice Information:")
        idn = device.send_command('*IDN?')
        print(idn)
        error = device.check_error()
        if error and error != "0":
            print(f"Error after IDN?: {error}")
        
        # Get current configuration
        print("\nCurrent Configuration:")
        mode = device.send_command('MODE?')
        print(f"Mode: {mode}")
        error = device.check_error()
        if error and error != "0":
            print(f"Error after MODE?: {error}")
        
        # Check first channel
        channel_info = device.send_command('CHNL?,1')
        print(f"Channel 1 Info: {channel_info}")
        error = device.check_error()
        if error and error != "0":
            print(f"Error after CHNL?: {error}")
        
        # Read voltages from all channels
        device.read_voltages()
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'device' in locals():
            device.disconnect()
            print("\nConnection closed")

if __name__ == "__main__":
    main()

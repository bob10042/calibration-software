import socket
import time
import sys

class APSM2000:
    def __init__(self, ip, port=10733):
        """Initialize connection to APS M2000"""
        self.ip = ip
        self.port = port
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.settimeout(3.0)  # 3 second timeout for commands

    def connect(self):
        """Establish connection to the device"""
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
        """Close the connection"""
        try:
            self.socket.close()
        except socket.error as e:
            print(f"Error closing connection: {e}")

    def send_command(self, cmd):
        """Send command and return response"""
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
        """Check for any errors"""
        return self.send_command('*ERR?')

    def format_reading(self, value, measurement_type):
        """Format scientific notation to readable value with proper units"""
        try:
            if not value:
                return "No reading"
            
            # Get first value before any spaces (main measurement)
            main_part = value.strip().split()[0]
            
            # Parse scientific notation
            if 'E' in main_part:
                mantissa, exponent = main_part.split('E')
                main_value = float(mantissa) * (10 ** int(exponent))
            else:
                main_value = float(main_part)
            
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
        """Read voltages from all channels"""
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

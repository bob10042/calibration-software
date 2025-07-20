import socket
import time

def check_computer_network():
    """
    Check if computer's network is properly configured.
    
    Returns:
        bool: True if network is properly configured, False otherwise
    """
    import subprocess
    import re
    
    try:
        print("\nChecking computer network configuration...")
        # Run ipconfig to get network info
        result = subprocess.run(['ipconfig'], capture_output=True, text=True)
        
        # Check if Ethernet adapter has correct IP
        if "192.168.15.101" not in result.stdout:
            print("\nComputer network not properly configured!")
            print("Please run one of these commands as administrator:")
            print("1. PowerShell: .\\set_network.ps1")
            print("2. Command Prompt: netsh interface ipv4 set address name=\"Ethernet\" static 192.168.15.101 255.255.255.0 192.168.15.254")
            print("\nAfter configuring network, wait 30 seconds then run this script again")
            return False
            
        print("✓ Computer network properly configured")
        return True
        
    except Exception as e:
        print(f"Error checking network configuration: {e}")
        return False

def validate_ip(ip):
    """
    Validate IP address format and connectivity.
    
    Args:
        ip (str): IP address to validate
        
    Returns:
        bool: True if IP is valid and reachable, False otherwise
    """
    import subprocess
    import re
    
    # Validate IP format
    ip_pattern = re.compile(r'^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')
    if not ip_pattern.match(ip):
        print(f"Invalid IP address format: {ip}")
        return False
    
    # Check if IP is reachable
    try:
        print(f"\nChecking network connectivity to {ip}...")
        result = subprocess.run(['ping', '-n', '1', '-w', '1000', ip], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print("Network connectivity confirmed")
            return True
        else:
            print(f"Unable to reach {ip}")
            print("\nPossible issues:")
            print("1. Device is not powered on")
            print("2. Device is not connected to the network")
            print("3. Device IP address is incorrect")
            print("4. Network cable is disconnected")
            return False
    except Exception as e:
        print(f"Error checking network connectivity: {e}")
        return False

def setup_lan_connection(ip_address, port=10733, local_ip=None):
    """
    Setup TCP/IP connection to APS M2000.
    
    Args:
        ip_address (str): IP address of the device
        port (int): Port number (default 10733 per APS manual)
        local_ip (str, optional): Local IP to bind to
        
    Returns:
        socket.socket: Connected socket object
    """
    try:
        # Create socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(180)  # 180 second timeout
        
        # Bind to specific local IP if provided
        if local_ip:
            sock.bind((local_ip, 0))
        
        # Connect to the device
        print(f"\nAttempting to connect to {ip_address}:{port}...")
        print("This may take a few seconds...")
        sock.connect((ip_address, port))
        print(f"TCP connection established successfully to {ip_address}")
        
        return sock
    except Exception as e:
        print(f"Error connecting to device: {e}")
        return None

def send_command(sock, command, wait_time=1.0):
    """
    Send a command to the APS M2000 and read the response if it's a query.
    
    Args:
        sock (socket.socket): Connected socket object
        command (str): Command to send
        wait_time (float): Time to wait before reading response
        
    Returns:
        str: Response from the device if command ends with ?, None otherwise
    """
    try:
        # Add newline termination if not present
        command = command.strip()
        if not command.endswith('\n'):
            command += '\n'
            
        print(f"Sending command: {command.strip()}")
        # Send command
        sock.send(command.encode())
        time.sleep(wait_time)  # Wait after every command
        
        # If it's a query, read response
        if command.strip().endswith('?'):
            response = sock.recv(4096).decode().strip()
            print(f"Received response: {response}")
            return response
        return None
    except Exception as e:
        print(f"Error sending command: {e}")
        return None

def initialize_device(sock):
    """
    Initialize the device following the measurement routine sequence.
    
    Args:
        sock (socket.socket): Connected socket object
        
    Returns:
        bool: True if initialization is successful, False otherwise
    """
    try:
        print("\nInitializing device...")
        
        # 1. Query device ID to verify communication
        print("1. Verifying device ID...")
        response = send_command(sock, "*IDN?", wait_time=2.0)
        if not response:
            print("Failed to get device identification")
            return False
        print("✓ Device identified:", response)
        
        # 2. Reset device to default state
        print("2. Resetting device...")
        send_command(sock, "*RST")
        time.sleep(2.0)  # Wait for reset to complete
        
        print("\nInitialization complete")
        return True
        
    except Exception as e:
        print(f"Error in device initialization: {e}")
        return False

def configure_channel(sock, channel, volt_range=100):
    """
    Configure a specific channel for voltage measurement.
    
    Args:
        sock (socket.socket): Connected socket object
        channel (int): Channel number (1-4)
        volt_range (float): Voltage range to set (default 100V)
    """
    try:
        # Select channel
        send_command(sock, f"CHAN {channel}")
        time.sleep(1.0)
        
        # Set voltage range
        send_command(sock, f"VOLT:RANGE {volt_range}")
        time.sleep(1.0)
        
        # Save configuration
        send_command(sock, "SAVECONFIG")
        time.sleep(1.0)
        
        print(f"Channel {channel} configured successfully")
        return True
        
    except Exception as e:
        print(f"Error configuring channel {channel}: {e}")
        return False

def query_voltage(sock, channel):
    """
    Query voltage measurement for a specific channel.
    
    Args:
        sock (socket.socket): Connected socket object
        channel (int): Channel number to query
        
    Returns:
        str: Voltage measurement value
    """
    try:
        # Select channel
        send_command(sock, f"CHAN {channel}")
        time.sleep(1.0)
        
        # Query voltage
        voltage = send_command(sock, "MEAS:VOLT?")
        if voltage:
            print(f"Channel {channel} voltage: {voltage.strip()}V")
        return voltage
        
    except Exception as e:
        print(f"Error querying voltage for channel {channel}: {e}")
        return None

def main():
    print("\nAPS M2000 Voltage Measurement Setup")
    print("===================================")
    
    # First check computer's network configuration
    if not check_computer_network():
        return
        
    # Device IP address
    ip_address = "192.168.15.100"
    print("\nAttempting to connect to device at:", ip_address)
    
    # Validate IP and connect
    if not validate_ip(ip_address):
        return
        
    sock = setup_lan_connection(ip_address=ip_address, port=10733)
    if not sock:
        return
            
    try:
        # Initialize the device
        if not initialize_device(sock):
            sock.close()
            return
            
        # Configure channels for voltage measurement
        for channel in range(1, 4):  # Channels 1-3
            print(f"\nConfiguring Channel {channel}...")
            if not configure_channel(sock, channel, volt_range=100):
                sock.close()
                return
            time.sleep(1.0)  # Wait between channel configurations
        
        # Query voltage measurements
        print("\nReading voltage measurements...")
        for channel in range(1, 4):
            voltage = query_voltage(sock, channel)
            if not voltage:
                print(f"Failed to read voltage on Channel {channel}")
        
        # Close connection
        sock.close()
        print("\nMeasurements complete - Connection closed")
        
    except Exception as e:
        print(f"\nError during measurement routine: {e}")
        if sock:
            sock.close()

if __name__ == "__main__":
    main()

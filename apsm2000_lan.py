import socket
import time

def setup_lan_connection(host='192.168.1.100', port=10733):
    """
    Sets up the LAN connection for communication with the APS M2000 Power Analyzer.

    Parameters:
        host (str): The IP address of the APS M2000 device.
        port (int): The port number for communication (default is 10733).

    Returns:
        socket.socket: Configured socket object.
    """
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((host, port))
        print(f"Connected to {host}:{port} successfully.")
        return sock
    except Exception as e:
        print(f"Error connecting to {host}:{port}: {e}")
        return None

def send_command(sock, command):
    """
    Sends a command to the APS M2000 and reads the response.

    Parameters:
        sock (socket.socket): Configured socket object.
        command (str): SCPI command to send.

    Returns:
        str: Response from the device or error message.
    """
    try:
        sock.sendall((command + '\n').encode())
        time.sleep(0.1)  # Allow some time for the device to respond
        response = sock.recv(1024).decode().strip()
        if response.startswith("ERR"):
            print(f"Device error: {response}")
        return response
    except Exception as e:
        print(f"Error sending command {command}: {e}")
        return None

def check_errors(sock):
    """
    Checks the device for any error messages.

    Parameters:
        sock (socket.socket): Configured socket object.

    Returns:
        list: List of error messages from the device.
    """
    errors = []
    try:
        response = send_command(sock, "SYST:ERR?")
        while response != "0,\"No error\"":
            errors.append(response)
            response = send_command(sock, "SYST:ERR?")
        return errors
    except Exception as e:
        print(f"Error checking device errors: {e}")
        return errors

def initialize_device(sock, mode="voltage"):
    """
    Initializes the APS M2000 device by sending necessary setup commands.

    Parameters:
        sock (socket.socket): Configured socket object.
        mode (str): Measurement mode, either "voltage" or "current".

    Returns:
        bool: True if initialization is successful, False otherwise.
    """
    try:
        # Clear any previous errors or states
        print("Clearing errors...")
        send_command(sock, "*CLS")

        # Reset the device
        print("Resetting device...")
        send_command(sock, "*RST")

        # Set the device to remote mode
        print("Setting to remote mode...")
        send_command(sock, "SYSTEM:REMOTE")

        # Configure measurement channels based on mode
        print(f"Configuring measurement channels for {mode}...")
        if mode == "voltage":
            send_command(sock, "MEAS:VOLTAGE:ACDC CH1")
            send_command(sock, "MEAS:VOLTAGE:ACDC CH2")
            send_command(sock, "MEAS:VOLTAGE:ACDC CH3")
        elif mode == "current":
            send_command(sock, "MEAS:CURRENT:ACDC CH1")
            send_command(sock, "MEAS:CURRENT:ACDC CH2")
            send_command(sock, "MEAS:CURRENT:ACDC CH3")
        else:
            print("Invalid mode specified.")
            return False

        # Set integration period (if needed)
        print("Setting integration period...")
        send_command(sock, "CONF:INTEGRATION 0.01")

        # Check if the device is ready
        print("Checking device status...")
        response = send_command(sock, "*OPC?")
        if response == "1":
            print("Device initialization successful.")
            return True
        else:
            print("Device initialization failed.")
            return False
    except Exception as e:
        print(f"Error during initialization: {e}")
        return False

def read_measurements(sock, channels, mode="voltage"):
    """
    Reads measurements (voltage or current) from specified channels.

    Parameters:
        sock (socket.socket): Configured socket object.
        channels (list): List of channel numbers to read from.
        mode (str): Measurement mode, either "voltage" or "current".

    Returns:
        dict: Channel numbers as keys and readings as values.
    """
    measurements = {}
    command_prefix = "MEAS:VOLTAGE:ACDC?" if mode == "voltage" else "MEAS:CURRENT:ACDC?"
    for channel in channels:
        command = f"{command_prefix} CH{channel}"
        response = send_command(sock, command)
        if response is not None:
            try:
                measurements[channel] = f"{float(response):.3f}"  # Format to 3 decimal places
            except ValueError:
                measurements[channel] = "Invalid Response"
        else:
            measurements[channel] = "Error"
    return measurements

def stream_measurements(sock, channels, mode="voltage"):
    """
    Continuously streams measurements (voltage or current) until the user terminates.

    Parameters:
        sock (socket.socket): Configured socket object.
        channels (list): List of channel numbers to read from.
        mode (str): Measurement mode, either "voltage" or "current".
    """
    try:
        print(f"Starting {mode} streaming. Press Ctrl+C to stop.")
        while True:
            measurement_readings = read_measurements(sock, channels, mode)
            print(f"\nMeasured {mode.capitalize()}s:")
            for channel, measurement in measurement_readings.items():
                print(f"Channel {channel}: {measurement} {'V' if mode == 'voltage' else 'A'}")
            time.sleep(1)  # Delay between readings
    except KeyboardInterrupt:
        print(f"\n{mode.capitalize()} streaming stopped by user.")

if __name__ == "__main__":
    # Configure LAN connection
    lan_socket = setup_lan_connection(host='192.168.15.100', port=10733)

    if lan_socket:
        # Choose mode (voltage or current)
        mode = input("Enter measurement mode (voltage/current): ").strip().lower()

        # Initialize the device
        if initialize_device(lan_socket, mode=mode):
            # List of channels to read from
            channels = [1, 2, 3]

            # Stream measurements continuously until the user stops
            stream_measurements(lan_socket, channels, mode=mode)

        # Close the LAN connection
        lan_socket.close()
        print("LAN connection closed.")

import serial
import time
import serial.tools.list_ports

def auto_select_serial_port():
    """
    Automatically selects the first available serial port.

    Returns:
        str: The name of the selected serial port, or None if no ports are available.
    """
    ports = serial.tools.list_ports.comports()
    if ports:
        print("Available serial ports:")
        for i, port in enumerate(ports):
            print(f"{i + 1}: {port.device}")
        selected_port = ports[0].device
        print(f"Auto-selected serial port: {selected_port}")
        return selected_port
    else:
        print("No serial ports available.")
        return None

def setup_serial_port(port='COM1', baudrate=115200, timeout=1):
    """
    Sets up the serial port for communication with the APS M2000 Power Analyzer.

    Parameters:
        port (str): The COM port to use (e.g., 'COM1').
        baudrate (int): Baud rate for serial communication.
        timeout (int): Read timeout in seconds.

    Returns:
        serial.Serial: Configured serial port object.
    """
    try:
        ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
            rtscts=True
        )
        print(f"Serial port {port} opened successfully.")
        return ser
    except Exception as e:
        print(f"Error opening serial port {port}: {e}")
        return None

def send_command(ser, command):
    """
    Sends a command to the APS M2000 and reads the response.

    Parameters:
        ser (serial.Serial): Configured serial port object.
        command (str): SCPI command to send.

    Returns:
        str: Response from the device or error message.
    """
    try:
        if ser and ser.is_open:
            ser.write((command + '\n').encode())
            time.sleep(0.1)  # Allow some time for the device to respond
            response = ser.readline().decode().strip()
            if response.startswith("ERR"):
                print(f"Device error: {response}")
            return response
        else:
            print("Serial port is not open.")
            return None
    except Exception as e:
        print(f"Error sending command {command}: {e}")
        return None

def check_errors(ser):
    """
    Checks the device for any error messages.

    Parameters:
        ser (serial.Serial): Configured serial port object.

    Returns:
        list: List of error messages from the device.
    """
    errors = []
    try:
        response = send_command(ser, "SYST:ERR?")
        while response != "0,\"No error\"":
            errors.append(response)
            response = send_command(ser, "SYST:ERR?")
        return errors
    except Exception as e:
        print(f"Error checking device errors: {e}")
        return errors

def initialize_device(ser, mode="voltage"):
    """
    Initializes the APS M2000 device by sending necessary setup commands.

    Parameters:
        ser (serial.Serial): Configured serial port object.
        mode (str): Measurement mode, either "voltage" or "current".

    Returns:
        bool: True if initialization is successful, False otherwise.
    """
    try:
        # Clear any previous errors or states
        print("Clearing errors...")
        send_command(ser, "*CLS")

        # Reset the device
        print("Resetting device...")
        send_command(ser, "*RST")

        # Set the device to remote mode
        print("Setting to remote mode...")
        send_command(ser, "SYSTEM:REMOTE")

        # Configure measurement channels based on mode
        print(f"Configuring measurement channels for {mode}...")
        if mode == "voltage":
            send_command(ser, "MEAS:VOLTAGE:ACDC CH1")
            send_command(ser, "MEAS:VOLTAGE:ACDC CH2")
            send_command(ser, "MEAS:VOLTAGE:ACDC CH3")
        elif mode == "current":
            send_command(ser, "MEAS:CURRENT:ACDC CH1")
            send_command(ser, "MEAS:CURRENT:ACDC CH2")
            send_command(ser, "MEAS:CURRENT:ACDC CH3")
        else:
            print("Invalid mode specified.")
            return False

        # Set integration period (if needed)
        print("Setting integration period...")
        send_command(ser, "CONF:INTEGRATION 0.01")

        # Check if the device is ready
        print("Checking device status...")
        response = send_command(ser, "*OPC?")
        if response == "1":
            print("Device initialization successful.")
            return True
        else:
            print("Device initialization failed.")
            return False
    except Exception as e:
        print(f"Error during initialization: {e}")
        return False

def read_measurements(ser, channels, mode="voltage"):
    """
    Reads measurements (voltage or current) from specified channels.

    Parameters:
        ser (serial.Serial): Configured serial port object.
        channels (list): List of channel numbers to read from.
        mode (str): Measurement mode, either "voltage" or "current".

    Returns:
        dict: Channel numbers as keys and readings as values.
    """
    measurements = {}
    command_prefix = "MEAS:VOLTAGE:ACDC?" if mode == "voltage" else "MEAS:CURRENT:ACDC?"
    for channel in channels:
        command = f"{command_prefix} CH{channel}"
        response = send_command(ser, command)
        if response is not None:
            try:
                measurements[channel] = f"{float(response):.3f}"  # Format to 3 decimal places
            except ValueError:
                measurements[channel] = "Invalid Response"
        else:
            measurements[channel] = "Error"
    return measurements

def stream_measurements(ser, channels, mode="voltage"):
    """
    Continuously streams measurements (voltage or current) until the user terminates.

    Parameters:
        ser (serial.Serial): Configured serial port object.
        channels (list): List of channel numbers to read from.
        mode (str): Measurement mode, either "voltage" or "current".
    """
    try:
        print(f"Starting {mode} streaming. Press Ctrl+C to stop.")
        while True:
            measurement_readings = read_measurements(ser, channels, mode)
            print(f"\nMeasured {mode.capitalize()}s:")
            for channel, measurement in measurement_readings.items():
                print(f"Channel {channel}: {measurement} {'V' if mode == 'voltage' else 'A'}")
            time.sleep(1)  # Delay between readings
    except KeyboardInterrupt:
        print(f"\n{mode.capitalize()} streaming stopped by user.")

if __name__ == "__main__":
    # Automatically select a serial port
    selected_port = auto_select_serial_port()

    if selected_port:
        # Configure serial port
        serial_port = setup_serial_port(port=selected_port, baudrate=115200, timeout=2)

        if serial_port:
            # Choose mode (voltage or current)
            mode = input("Enter measurement mode (voltage/current): ").strip().lower()

            # Initialize the device
            if initialize_device(serial_port, mode=mode):
                # List of channels to read from
                channels = [1, 2, 3]

                # Stream measurements continuously until the user stops
                stream_measurements(serial_port, channels, mode=mode)

            # Close the serial port
            serial_port.close()
            print("Serial port closed.")

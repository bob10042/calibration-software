import serial
import time
import serial.tools.list_ports

def select_serial_port():
    ports = list(serial.tools.list_ports.comports())
    if ports:
        print("\nAvailable serial ports:")
        for i, port in enumerate(ports):
            print(f"{i + 1}: {port.device} - {port.description}")
        
        while True:
            try:
                choice = input("\nSelect port number (or press Ctrl+C to exit): ")
                port_num = int(choice) - 1
                if 0 <= port_num < len(ports):
                    selected_port = ports[port_num].device
                    print(f"\nSelected port: {selected_port}")
                    return selected_port
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Please enter a number.")
    else:
        print("No serial ports available.")
        return None

def setup_serial_port(port='COM1', baudrate=115200, timeout=1, max_retries=3):
    for attempt in range(max_retries):
        try:
            print(f"\nAttempt {attempt + 1} of {max_retries} to open {port}")
            
            try:
                print("Checking if port is in use...")
                temp_ser = serial.Serial(port)
                print("Port was open, closing it...")
                temp_ser.close()
                temp_ser = None
                time.sleep(2)
            except Exception as e:
                print(f"Port check result: {str(e)}")
            
            print(f"Attempting to open {port} at {baudrate} baud...")
            ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=timeout,
                rtscts=False,
                write_timeout=2,
                exclusive=True  # Windows requires exclusive access
            )
            
            # Set DTR and wait
            ser.setDTR(True)
            time.sleep(0.5)
            
            if ser.is_open:
                print(f"Successfully opened {port}")
                # Test the connection
                ser.write(b'\r\n')  # Send a blank line
                time.sleep(0.1)
                ser.reset_input_buffer()  # Clear any response
                return ser
            
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {str(e)}")
            time.sleep(2)  # Wait before retry
            
    print(f"\nFailed to open {port} after {max_retries} attempts")
    return None

def send_command(ser, command, timeout=1.0):
    try:
        if ser and ser.is_open:
            # Clear buffers
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            
            print(f"\nSending command: {command}")
            # Add newline only (no carriage return)
            full_command = command + '\n'
            print(f"Sending bytes: {[hex(b) for b in full_command.encode()]}")
            
            # Send command
            ser.write(full_command.encode())
            ser.flush()
            time.sleep(0.2)  # Short delay after sending
            
            # Read with timeout
            start_time = time.time()
            response = ''
            
            while (time.time() - start_time) < timeout:
                if ser.in_waiting:
                    new_data = ser.read(ser.in_waiting).decode(errors='ignore')
                    print(f"Received bytes: {[hex(b) for b in new_data.encode()]}")
                    response += new_data
                    if '\n' in new_data:  # Got a complete line
                        break
                time.sleep(0.1)
            
            response = response.strip()
            print(f"Final response: {repr(response)}")
            
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
    try:
        # Clear any previous errors or states
        print("Clearing errors...")
        send_command(ser, "*CLS")

        # Reset the device
        print("Resetting device...")
        send_command(ser, "*RST")

        # Set the device to remote mode
        print("Setting to remote mode...")
        send_command(ser, "REMOTE")

        # Configure measurement channels based on mode
        print(f"Configuring measurement channels for {mode}...")
        if mode == "voltage":
            for channel in [1, 2, 3]:
                send_command(ser, f"CHAN {channel}")
                send_command(ser, "VOLT:RANGE 100")
                send_command(ser, "SAVECONFIG")
        elif mode == "current":
            for channel in [1, 2, 3]:
                send_command(ser, f"CHAN {channel}")
                send_command(ser, "CURR:RANGE 10")
                send_command(ser, "SAVECONFIG")
        else:
            print("Invalid mode specified.")
            return False

        # Set integration period
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
    measurements = {}
    command = "MEAS:VOLT?" if mode == "voltage" else "MEAS:CURR?"
    for channel in channels:
        send_command(ser, f"CHAN {channel}")
        time.sleep(0.1)  # Give device time to switch channels
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
    try:
        print(f"Starting {mode} streaming. Press Ctrl+C to stop.")
        while True:
            measurement_readings = read_measurements(ser, channels, mode)
            print(f"\nMeasured {mode.capitalize()}s:")
            for channel, measurement in measurement_readings.items():
                print(f"Channel {channel}: {measurement} {'V' if mode == 'voltage' else 'A'}")
            time.sleep(1)
    except KeyboardInterrupt:
        print(f"\n{mode.capitalize()} streaming stopped by user.")

def test_device_communication(ser):
    """Simple test function to verify device communication"""
    try:
        print("\nTesting device communication...")
        
        # Clear any pending data
        ser.reset_input_buffer()
        ser.reset_output_buffer()
        
        print("\nSending initial test command...")
        command_bytes = b'\r\n'  # Just send a blank line first
        ser.write(command_bytes)
        ser.flush()
        time.sleep(1)
        
        print("\nSending *IDN? command...")
        command_bytes = b'*IDN?\r\n'
        print("Sending bytes:", [hex(b) for b in command_bytes])
        ser.write(command_bytes)
        ser.flush()
        
        # Wait up to 5 seconds for a response
        print("Waiting for response...")
        start_time = time.time()
        response = ''
        while (time.time() - start_time) < 5:
            if ser.in_waiting:
                new_bytes = ser.read(ser.in_waiting)
                print(f"Received bytes: {[hex(b) for b in new_bytes]}")
                response += new_bytes.decode(errors='ignore')
                if '\n' in response:
                    break
            time.sleep(0.1)
        
        print(f"Final response: {repr(response)}")
        input("Press Enter to continue or Ctrl+C to exit...")

        print("\n2. Checking for errors")
        response = send_command(ser, "SYST:ERR?")
        input(f"Error check response: {response}\nPress Enter to continue or Ctrl+C to exit...")

        print("\n3. Attempting to set remote mode")
        response = send_command(ser, "REMOTE")
        time.sleep(0.5)  # Give device time to switch to remote mode
        input("Please check if the device shows 'Remote' on its display.\nPress Enter if it does, or Ctrl+C to exit...")

        return True
    except Exception as e:
        print(f"Error during communication test: {e}")
        return False

if __name__ == "__main__":
    selected_port = select_serial_port()

    if selected_port:
        print("\nWaiting for device to be ready...")
        time.sleep(3)  # Wait for device to fully enumerate
        # Try different baud rates
        for baudrate in [115200, 19200, 9600]:
            print(f"\nTrying baud rate: {baudrate}")
            serial_port = setup_serial_port(port=selected_port, baudrate=baudrate, timeout=2)
            if serial_port:
                break

        if serial_port:
            if test_device_communication(serial_port):
                print("\nDevice communication test successful!")
            else:
                print("\nDevice communication test failed!")

            serial_port.close()
            print("Serial port closed.")

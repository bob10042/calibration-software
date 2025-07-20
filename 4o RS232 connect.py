import serial
import time
import logging
import serial.tools.list_ports

# Logging setup
logging.basicConfig(level=logging.DEBUG, filename="apsm2000_rs232_debug.log", filemode="w",
                    format="%(asctime)s - %(levelname)s - %(message)s")

def list_available_ports():
    """List all available serial ports."""
    ports = list(serial.tools.list_ports.comports())
    if ports:
        logging.info("Available serial ports:")
        for port in ports:
            logging.info(f"- {port.device}")
        return ports[0].device  # Return first available port
    logging.error("No serial ports available")
    return None

def open_serial_connection(port, baudrate=9600, timeout=2):
    """Open an RS232 connection to the device."""
    try:
        ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
            rtscts=False  # Disable hardware flow control as per documentation
        )
        logging.info(f"Opened serial connection on {port} with baudrate {baudrate}.")
        return ser
    except Exception as e:
        logging.error(f"Failed to open serial connection: {e}")
        return None

def send_command(ser, command):
    """Send a command over RS232."""
    try:
        command += "\n"  # Use \n termination as per documentation
        ser.write(command.encode())
        logging.info(f"Sent command: {command.strip()}.")
        time.sleep(0.1)  # Small delay after sending command
    except Exception as e:
        logging.error(f"Failed to send command '{command.strip()}': {e}")

def read_response(ser, max_attempts=5):
    """Read the response from the device."""
    try:
        for attempt in range(max_attempts):
            response = ser.readline().decode(errors='replace').strip()
            if response:
                logging.info(f"Received response: {response}.")
                return response
            logging.debug(f"Attempt {attempt + 1}: No response yet, retrying...")
            time.sleep(0.5)
        logging.warning("No response received after max attempts.")
        return None
    except Exception as e:
        logging.error(f"Failed to read response: {e}")
        return None

def initialize_device(ser):
    """Initialize the device and query IDN."""
    try:
        # Clear and reset
        send_command(ser, "*CLS")
        time.sleep(0.5)
        send_command(ser, "*RST")
        time.sleep(1)

        # Enter remote mode using correct command
        send_command(ser, "REMOTE")
        time.sleep(0.5)

        # Query IDN
        send_command(ser, "*IDN?")
        idn_response = read_response(ser)
        if idn_response:
            print(f"Device Identification: {idn_response}")
            logging.info(f"Device Identification: {idn_response}")
        else:
            logging.warning("*IDN? command failed. Checking for errors...")
            send_command(ser, "SYST:ERR?")
            error_response = read_response(ser)
            if error_response:
                logging.error(f"Error Response: {error_response}")
            else:
                logging.info("No errors reported by the device.")
    except Exception as e:
        logging.error(f"Error during initialization: {e}")

def main():
    # Auto-detect available ports
    port = list_available_ports()
    if not port:
        print("No serial ports available. Please check your connections.")
        return

    print(f"Attempting to connect to {port}")
    # Try with 9600 baud rate first as it's the most common default
    ser = open_serial_connection(port, baudrate=9600)

    if not ser:
        print("Failed to open serial connection. Check logs for details.")
        return

    try:
        initialize_device(ser)
    finally:
        ser.close()
        logging.info("Closed serial connection.")
        print("Closed serial connection.")

if __name__ == "__main__":
    main()

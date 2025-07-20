import socket
import logging
import time

# Logging setup
logging.basicConfig(level=logging.DEBUG, filename="apsm2000_lan_debug.log", filemode="w",
                    format="%(asctime)s - %(levelname)s - %(message)s")

def setup_lan_connection(host='192.168.15.100', port=10733):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((host, port))
        print(f"Connected to {host}:{port} successfully.")
        return sock
    except Exception as e:
        print(f"Error connecting to {host}:{port}: {e}")
        logging.error(f"Error connecting to {host}:{port}: {e}")
        return None

def send_command(sock, command, expect_response=True):
    try:
        # Add carriage return and newline
        sock.sendall((command + '\r\n').encode())
        time.sleep(0.5)
        if expect_response:
            response = sock.recv(1024).decode().strip()
            logging.info(f"Command: {command}, Response: {response}")
            return response
        return None
    except Exception as e:
        logging.error(f"Error sending command {command}: {e}")
        return None

def initialize_device(sock):
    try:
        # Query IDN first
        print("Querying device identification...")
        idn_response = send_command(sock, "*IDN?")
        if idn_response:
            print(f"Device Identification: {idn_response}")
        else:
            print("Failed to get device identification.")
            return False

        # Set the device to remote mode
        print("Setting to remote mode...")
        send_command(sock, "REMOTE", expect_response=False)
        
        # Configure channels
        print("Configuring channels...")
        for channel in range(1, 4):
            print(f"Setting up channel {channel}...")
            send_command(sock, f"CHAN {channel}", expect_response=False)
            send_command(sock, "VOLT:RANGE 100", expect_response=False)
            send_command(sock, "SAVECONFIG", expect_response=False)
            time.sleep(0.5)
        
        # Start data logging
        print("Starting data logging...")
        send_command(sock, "DATALOG 1", expect_response=False)
        
        return True
    except Exception as e:
        print(f"Error during initialization: {e}")
        logging.error(f"Error during initialization: {e}")
        return False

def stream_voltage(sock):
    try:
        print("Starting voltage streaming. Press Ctrl+C to stop.")
        while True:
            for channel in range(1, 4):
                # Select channel first
                send_command(sock, f"CHAN {channel}", expect_response=False)
                # Then measure voltage
                response = send_command(sock, "MEAS:VOLT?")
                if response:
                    print(f"Channel {channel}: {response} V")
                else:
                    print(f"Channel {channel}: Error reading voltage.")
            time.sleep(1)  # Delay between readings
    except KeyboardInterrupt:
        print("Voltage streaming stopped by user.")
    except Exception as e:
        print(f"Error during voltage streaming: {e}")
        logging.error(f"Error during voltage streaming: {e}")

if __name__ == "__main__":
    host = "192.168.15.100"  # Correct IP address
    port = 10733  # Static port for LAN communication

    sock = setup_lan_connection(host, port)

    if sock:
        if initialize_device(sock):
            stream_voltage(sock)
        sock.close()
        print("LAN connection closed.")

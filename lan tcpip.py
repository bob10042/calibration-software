import socket
import time

def read_voltages_via_tcp(ip_address="192.168.1.100", port=10733):
    """
    Example of opening a TCP/IP socket to the M2000,
    sending a command to read all channel voltages,
    and printing the response.
    """
    # Create a TCP socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        # Optional: set a timeout on blocking socket operations
        s.settimeout(2.0)  # seconds

        print(f"Connecting to {ip_address}:{port} ...")
        s.connect((ip_address, port))
        print("Connected.")

        # Clear interface and errors:
        s.sendall(b"*CLS\n")
        time.sleep(0.2)

        # Send the READ? command
        command = "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3, VOLTS:CH4\n"
        s.sendall(command.encode("ascii"))

        # The M2000 typically ends responses with \r\n
        # We'll do a simple recv loop until we get the \n
        response_buffer = []
        while True:
            chunk = s.recv(4096).decode("ascii")
            if not chunk:
                # Connection closed or no data
                break
            response_buffer.append(chunk)
            # If we see a newline, we assume that's the end
            if "\n" in chunk:
                break

        raw_response = "".join(response_buffer).strip()
        print("Raw response:", raw_response)

        # Split on commas
        if raw_response:
            voltage_strings = raw_response.split(',')
            voltages = [float(v) for v in voltage_strings]
            print("Parsed channel voltages:", voltages)
        else:
            print("No response or empty response received.")

if __name__ == "__main__":
    read_voltages_via_tcp("192.168.1.100", 10733)

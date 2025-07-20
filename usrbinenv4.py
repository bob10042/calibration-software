import socket
import time

def init_and_read_idn_lan(ip_address="192.168.1.100", port=10733, lock_front_panel=False):
    """
    Opens a TCP socket to the APSM2000 at IP:port.
    1) Clears interface.
    2) (Optionally) LOCKOUT the front panel.
    3) Requests *IDN?.
    4) Prints ID string.
    5) Closes connection.
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(2.0)  # 2 second timeout
    try:
        print(f"Connecting to {ip_address}:{port} ...")
        s.connect((ip_address, port))
        print("Connected.")

        # Clear interface
        s.sendall(b"*CLS\n")
        time.sleep(0.1)

        if lock_front_panel:
            s.sendall(b"LOCKOUT\n")
            time.sleep(0.1)

        # Request ID
        s.sendall(b"*IDN?\n")
        
        response_buffer = []
        while True:
            chunk = s.recv(4096)
            if not chunk:
                break
            response_buffer.append(chunk)
            if b"\n" in chunk:
                # Found the newline terminator
                break

        raw_response = b"".join(response_buffer).decode('ascii').strip()
        print("Raw *IDN? response:", raw_response)

        fields = raw_response.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Example: read 3-channel voltages
        # s.sendall(b"READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3\n")
        # resp_data = s.recv(4096).decode('ascii').strip()
        # print("3-ch voltage response:", resp_data)

        # Optionally revert to local:
        # s.sendall(b"LOCAL\n")

    except socket.timeout:
        print("Socket timed out.")
    except Exception as ex:
        print("LAN Error:", ex)
    finally:
        s.close()
        print("Socket closed.")

if __name__ == "__main__":
    init_and_read_idn_lan("192.168.1.100", 10733, lock_front_panel=False)

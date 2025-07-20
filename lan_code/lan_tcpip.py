import socket
import time

def send_command(sock, command, expect_response=False):
    """Helper function to send a command and get response if needed"""
    print(f"Sending: {command}")
    sock.sendall((command + "\n").encode())
    time.sleep(0.1)  # Same delay as working script
    
    if expect_response:
        try:
            # Use same buffer size as working script
            response = sock.recv(1024).decode().strip()
            if response:
                print(f"Response: {response}")
                return response
            if response.startswith("ERR"):
                print(f"Device error: {response}")
        except Exception as e:
            print(f"Error: {e}")
    return None

def read_voltages_via_tcp(ip_address="192.168.1.100", port=10733):
    """Read voltages from the M2000 power analyzer"""
    try:
        # Create socket with same settings as working script
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2.0)
        
        print(f"Connecting to {ip_address}:{port}...")
        sock.connect((ip_address, port))
        print("Connected successfully")
        
        with sock:
            # Initialize exactly like working script
            print("\nInitializing device...")
            
            # Clear any previous errors or states
            send_command(sock, "*CLS")
            
            # Reset the device
            send_command(sock, "*RST")
            time.sleep(0.1)
            
            # Set to remote mode
            send_command(sock, "SYSTEM:REMOTE")
            time.sleep(0.1)
            
            # Configure measurement channels
            print("\nConfiguring channels...")
            for ch in range(1, 5):
                send_command(sock, f"MEAS:VOLTAGE:ACDC CH{ch}")
                time.sleep(0.1)
            
            # Set integration period as in working script
            send_command(sock, "CONF:INTEGRATION 0.01")
            time.sleep(0.1)
            
            # Read voltages
            print("\nReading voltages...")
            voltages = {}
            for ch in range(1, 5):
                response = send_command(sock, f"MEAS:VOLTAGE:ACDC? CH{ch}", True)
                if response:
                    try:
                        voltages[f"CH{ch}"] = float(response)
                    except ValueError:
                        voltages[f"CH{ch}"] = f"Error: {response}"
                else:
                    voltages[f"CH{ch}"] = "No response"
                time.sleep(0.1)
            
            print("\nVoltage Readings:")
            for channel, value in voltages.items():
                print(f"{channel}: {value}")
                
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    read_voltages_via_tcp("192.168.15.100", 10733)
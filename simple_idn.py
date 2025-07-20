import serial
import time
import serial.tools.list_ports

def list_ports():
    ports = list(serial.tools.list_ports.comports())
    if ports:
        print("\nAvailable ports:")
        for i, port in enumerate(ports):
            print(f"{i + 1}: {port.device} - {port.description}")
        return ports
    return None

def try_idn(port, baudrate=115200):
    try:
        print(f"\nTrying {port} at {baudrate} baud...")
        ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=2,
            rtscts=True  # Enable hardware flow control
        )
        
        if ser.is_open:
            print("Port opened successfully")
            
            # Clear any pending data
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            
            # Send *IDN? command
            print("Sending *IDN?...")
            ser.write(b'*IDN?\n')
            time.sleep(0.1)  # Short delay
            
            # Read response using readline
            response = ser.readline().decode().strip()
            print(f"Response: {response}")
            
            ser.close()
            return True
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

if __name__ == "__main__":
    ports = list_ports()
    if not ports:
        print("No ports found")
        exit()
    
    while True:
        try:
            choice = input("\nSelect port number (or press Ctrl+C to exit): ")
            port_num = int(choice) - 1
            if 0 <= port_num < len(ports):
                port = ports[port_num].device
                print(f"\nSelected: {port}")
                
                # Try with 115200 baud only
                if try_idn(port):
                    print("\nCommunication successful!")
                break
            else:
                print("Invalid selection")
        except ValueError:
            print("Please enter a number")
        except KeyboardInterrupt:
            print("\nExiting...")
            break

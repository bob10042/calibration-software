import serial
import time

def send_command(ser, command):
    """Send a command and print the raw bytes for debugging"""
    if not command.endswith('\n'):
        command += '\n'
    
    try:
        # Convert command to bytes and show what we're sending
        cmd_bytes = command.encode()
        print(f"Sending command: {command.strip()}")
        print(f"Sending bytes: {cmd_bytes}")
        
        # Send command
        ser.write(cmd_bytes)
        ser.flush()
        time.sleep(1)  # Wait after sending
        
        # If it's a query, read response
        if '?' in command:
            # Wait longer for response
            time.sleep(1)
            
            # Try to read multiple times
            for _ in range(3):  # Try up to 3 times
                if ser.in_waiting:
                    response = ser.read(ser.in_waiting)
                    print(f"Received bytes: {response}")
                    print(f"Decoded response: {response.decode(errors='ignore')}")
                    return response.decode(errors='ignore').strip()
                time.sleep(1)  # Longer wait between tries
            
            print("No response received")
        return None
        
    except Exception as e:
        print(f"Error sending command: {str(e)}")
        return None

def main():
    try:
        # Open COM4 at 115200 baud with longer timeouts
        print("\nOpening COM4...")
        ser = serial.Serial(
            port='COM4',
            baudrate=115200,
            bytesize=8,
            parity='N',
            stopbits=1,
            timeout=2,          # Read timeout
            write_timeout=2,    # Write timeout
            xonxoff=False,      # No software flow control
            rtscts=False,       # No hardware flow control
            dsrdtr=False        # No hardware flow control
        )
        
        if not ser.is_open:
            ser.open()
        
        print("Port opened successfully")
        
        # Clear any existing data
        ser.reset_input_buffer()
        ser.reset_output_buffer()
        
        # Initialize the device step by step
        print("\nInitializing device...")
        
        # Step 1: Enter remote mode first
        print("\nStep 1: Enter remote mode")
        send_command(ser, "REMOTE")
        time.sleep(2)  # Give it time to enter remote mode
        
        # Step 2: Clear state
        print("\nStep 2: Clear state")
        send_command(ser, "*CLS")
        time.sleep(1)
        
        # Step 3: Reset device
        print("\nStep 3: Reset device")
        send_command(ser, "*RST")
        time.sleep(2)  # Give it time to reset
        
        # Step 4: Query ID
        print("\nStep 4: Query device ID")
        response = send_command(ser, "*IDN?")
        
        if response:
            print(f"\nDevice identified: {response}")
            
            # Try configuring a channel
            print("\nConfiguring channel 1...")
            commands = [
                "CHAN 1",            # Select channel 1
                "VOLT:RANGE 100",    # Set voltage range to 100V
                "SAVECONFIG",        # Save configuration
                "DATALOG 1",         # Enable data logging
                "MEAS:VOLT?"         # Try a measurement
            ]
            
            for cmd in commands:
                print(f"\nSending: {cmd}")
                response = send_command(ser, cmd)
                time.sleep(1)  # Wait between commands
                
        else:
            print("\nFailed to get device identification")
            print("Please verify:")
            print("1. Device is powered on")
            print("2. RS232 cable is properly connected")
            print("3. Device is set to 115200 baud")
            print("4. Check device's front panel for any error messages")
        
        # Close the port
        ser.close()
        print("\nPort closed")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        try:
            ser.close()
        except:
            pass

if __name__ == "__main__":
    main()

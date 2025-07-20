import serial
import time

def send_command(ser, cmd):
    """Send command and read response"""
    print(f"Sending: {cmd.strip()}")
    ser.write(f"{cmd}\n".encode())
    time.sleep(0.1)
    
    if '?' in cmd:
        response = ser.readline().decode().strip()
        print(f"Response: '{response}'")
        return response
    return None

def init_m2000():
    ser = serial.Serial(
        port='COM5',
        baudrate=115200,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        rtscts=True,
        timeout=1
    )
    
    # Basic initialization
    send_command(ser, '*CLS')
    send_command(ser, '*RST')
    time.sleep(0.5)  # Wait longer for reset
    
    # Verify communication
    idn = send_command(ser, '*IDN?')
    print(f"Connected to: {idn}")
    
    # Basic configuration
    commands = [
        'EDITCONFIG',
        'MODE,0',                  # Single VPA mode
        'SAVECONFIG'
    ]
    
    for cmd in commands:
        send_command(ser, cmd)
        time.sleep(0.1)
    
    return ser

def read_measurements(ser):
    # Following exact RDEF format from manual:
    # Measurement_Data:Measurement_Source:Measurement_Type
    
    read_formats = [
        # Format 1: Basic voltage reading
        'READ?VOLTS:1:COUPLED',
        
        # Format 2: Try with specific channel definition
        'READ?V:CH1:COUPLED',
        
        # Format 3: Basic format with VPA specification
        'READ?V:VPA1:COUPLED',
        
        # Format 4: With explicit coupling
        'READ?VOLTS:CH1:ACDC:COUPLED'
    ]
    
    for fmt in formats:
        print(f"\nTrying format: {fmt}")
        volts = send_command(ser, fmt)
        err = send_command(ser, '*ERR?')
        
        if err == '0' and volts:
            return volts
            
        # If failed, check MCR status
        mcr = send_command(ser, 'MCR?')
        print(f"MCR Status: {mcr}")
        
        time.sleep(0.2)  # Longer delay between attempts
    
    return None

def main():
    try:
        ser = init_m2000()
        print("\nInitialization complete")
        time.sleep(1)
        
        print("\nStarting measurements...")
        for i in range(3):
            print(f"\nReading {i+1}:")
            reading = read_measurements(ser)
            if reading:
                try:
                    value = float(reading)
                    print(f"Measurement: {value:.6f}")
                except:
                    print(f"Raw response: {reading}")
            else:
                print("No valid reading obtained")
            time.sleep(1)
            
    except serial.SerialException as e:
        print(f"Serial port error: {e}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print("\nSerial port closed")

if __name__ == "__main__":
    main()
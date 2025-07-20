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
    time.sleep(0.5)
    
    # Verify communication
    idn = send_command(ser, '*IDN?')
    print(f"Connected to: {idn}")
    
    # Configure the device
    commands = [
        'EDITCONFIG',
        'MODE,0',                  # Single VPA mode
        'CHANNELS,1,1',            # Add CH1 to VPA1
        'COUPLE,1,0',              # Set AC+DC coupling for VPA1
        'PERIOD,1,4',              # Set measurement period
        'SAVECONFIG'
    ]
    
    for cmd in commands:
        send_command(ser, cmd)
        time.sleep(0.1)
    
    return ser

def read_measurements(ser):
    # Check if measurement is ready
    mcr = send_command(ser, 'MCR?')
    if mcr != '1':
        print("Waiting for measurement completion...")
        time.sleep(0.5)
    
    # Try different READ formats from manual examples on page 271
    read_formats = [
        'READ?VOLTS:CH1',          # Basic channel voltage
        'READ?V,CH1',              # Alternative format
        'READ?FREQ,CH1',           # Try reading frequency
        'READ?VOLTS:VPA1',         # Try reading VPA voltage
        'READ?V,1',                # Using numeric channel
        'READ?VOLTS,CH1',          # Alternative separator
    ]
    
    for fmt in read_formats:
        print(f"\nTrying format: {fmt}")
        reading = send_command(ser, fmt)
        err = send_command(ser, '*ERR?')
        
        if err == '0' and reading:
            print(f"Success with format: {fmt}")
            return reading
        
        print(f"Error {err} with format: {fmt}")
        time.sleep(0.2)
    
    return None

def main():
    try:
        ser = init_m2000()
        print("\nInitialization complete")
        time.sleep(1)
        
        # Try to get VPA configuration first
        vpa_check = send_command(ser, 'VPA?1')
        print(f"VPA1 configuration: {vpa_check}")
        
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
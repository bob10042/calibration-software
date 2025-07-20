import serial
import time

def send_command(ser, cmd):
    """Send command and read response"""
    print(f"Sending: {cmd.strip()}")
    ser.write(f"{cmd}\n".encode())
    time.sleep(0.1)
    
    if '?' in cmd:  # If it's a query, read response
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
    
    # Clear any pending operations and reset
    send_command(ser, '*CLS')
    send_command(ser, '*RST')
    
    # Verify communication
    idn = send_command(ser, '*IDN?')
    if not idn:
        raise Exception("No response from device")
    print(f"Connected to: {idn}")
    
    # Configure VPA and channels
    commands = [
        'EDITCONFIG',
        'MODE,0',                  # Single VPA mode
        'CHANNELS,1,1',            # Assign channel 1 to VPA1
        'COUPLE,1,0',              # AC+DC coupling for VPA1
        'PERIOD,1,4',              # Set measurement period
        'BANDWIDTH,1,1',           # No bandwidth limiting
        'RESPONSE,1,1',            # Medium response
        'SAVECONFIG'               # Save configuration
    ]
    
    for cmd in commands:
        send_command(ser, cmd)
        time.sleep(0.1)
    
    # Verify configuration
    send_command(ser, 'VPA?1')     # Check channel 1 VPA assignment
    send_command(ser, 'COUPLE?1')   # Check coupling
    
    return ser

def read_measurements(ser):
    # Check measurement completion
    mcr = send_command(ser, 'MCR?')
    
    # Read channel 1 voltage using COUPLED type
    cmd = 'READ?VOLTS:CH1,COUPLED'
    volts = send_command(ser, cmd)
    
    # Check for errors
    err = send_command(ser, '*ERR?')
    if err != '0':
        print(f"Error after reading: {err}")
        
        # Try alternative format if first fails
        cmd = 'READ?V:CH1'
        volts = send_command(ser, cmd)
        err = send_command(ser, '*ERR?')
        if err != '0':
            print(f"Error with alternative format: {err}")
    
    return volts

def main():
    try:
        ser = init_m2000()
        print("\nInitialization complete")
        time.sleep(1)
        
        print("\nStarting measurements...")
        for i in range(3):
            print(f"\nReading {i+1}:")
            voltage = read_measurements(ser)
            if voltage:
                try:
                    volt_val = float(voltage)
                    print(f"Channel 1 Voltage (V): {volt_val:.6f}")
                except:
                    print(f"Raw voltage response: {voltage}")
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
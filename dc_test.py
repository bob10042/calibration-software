import serial
import time

def send_command(ser, cmd, max_retries=3):
    """Send command and get response with retry logic"""
    print(f"\nSending: {cmd}")
    
    for attempt in range(max_retries):
        try:
            # Clear buffers before sending
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            
            # Format command with proper termination
            cmd_str = f"{cmd}\n"
            
            # Send command
            ser.write(cmd_str.encode())
            time.sleep(1)  # Wait for command processing
            
            # Read response if it's a query
            if cmd.endswith('?'):
                response = ""
                # Read until we get a response or timeout
                start_time = time.time()
                while not response and (time.time() - start_time) < 2:
                    response = ser.read_all().decode().strip()
                    if not response:
                        time.sleep(0.1)
                
                print(f"Response: {response}")
                return response
            
            return ""
            
        except Exception as e:
            print(f"Command error: {str(e)}")
            if attempt < max_retries - 1:
                print(f"Attempt {attempt + 1}: Retrying...")
                time.sleep(1)
                continue
            raise
    
    return ""

def setup_dc_mode(ser):
    """Setup AGX for DC mode operation"""
    print("\nSetting up DC mode...")
    
    # Initial reset and clear
    send_command(ser, "*RST")
    time.sleep(2)
    send_command(ser, "*CLS")
    time.sleep(1)
    
    # First command sequence
    send_command(ser, ":OUTP,OFF;:INIT,OFF;:*LLO,")
    time.sleep(1)
    
    # Basic DC mode setup matching 3150AFX reference
    dc_setup = (":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,3;"
               ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;"
               ":VOLT:DC:LIM:MIN,0;:VOLT:DC:LIM:MAX,425;:COUPL,DIRECT;"
               ":PHAS2,120;:PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON")
    send_command(ser, dc_setup)
    time.sleep(2)
    
    # Verify DC mode is active
    mode = send_command(ser, "VOLT:MODE?")
    if mode and "DC" not in mode.upper():
        print("Failed to enter DC mode")
        return False
        
    return True

def measure_dc_voltage(ser):
    """Take DC voltage measurements"""
    total = 0.0
    for i in range(10):
        response = send_command(ser, ":MEAS:VOLT:DC?")
        if response:
            try:
                value = float(response.strip())
                total += value
                print(f"Reading {i+1}: {value:.3f}V DC")
            except ValueError:
                print(f"Invalid reading {i+1}")
    return total / 10

def set_dc_voltage(ser, voltage):
    """Set DC voltage and verify"""
    try:
        if not 0 <= voltage <= 425:
            print(f"Warning: Voltage {voltage}V outside valid range (0-425V)")
            voltage = min(max(voltage, 0), 425)
            print(f"Adjusting to {voltage}V")
        
        print(f"\nSetting DC voltage to {voltage}V...")
        
        # Disable output before voltage change
        send_command(ser, ":OUTP,OFF")
        time.sleep(1)
        
        # Set voltage using reference command format
        send_command(ser, f":VOLT,{voltage}")
        time.sleep(1)
        
        # Enable output
        send_command(ser, ":OUTP,ON")
        time.sleep(30)  # Match C# timing for initial stabilization
        
        # Verify output is enabled
        output = send_command(ser, "OUTP?")
        if not output or output.strip() != "1":
            print("Failed to enable output")
            return False
            
        # Take measurements
        print("\nMeasuring DC voltage:")
        v = measure_dc_voltage(ser)
        print(f"\nAverage: {v:.3f}V DC")
                
        return True
        
    except Exception as e:
        print(f"Error setting voltage: {str(e)}")
        return False
    finally:
        # Always ensure output is off after test
        send_command(ser, ":OUTP,OFF")

def main():
    ser = None
    try:
        # Connect to AGX
        ser = serial.Serial(
            port='COM3',
            baudrate=9600,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=5
        )
        print("Connected to COM3")
        
        # DC test sequence
        print("\n=== DC Voltage Test ===")
        if setup_dc_mode(ser):
            test_voltages = [10, 50, 100, 200, 300, 400]  # Test points up to 400V
            for voltage in test_voltages:
                if not set_dc_voltage(ser, voltage):
                    print(f"Failed at {voltage}V test point")
                    break
                time.sleep(15)  # Wait between test points
            
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        if ser:
            try:
                # Safe shutdown sequence matching C# implementation
                send_command(ser, ":OUTP,OFF")
                send_command(ser, "*GTL")
                send_command(ser, "*RST")
                ser.close()
                print("Device safely shut down")
            except:
                pass

if __name__ == "__main__":
    main()

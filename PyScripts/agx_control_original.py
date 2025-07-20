import serial
import time

def check_errors(ser):
    """Read and display all errors in the queue"""
    while True:
        ser.write("SYST:ERR?\n".encode())
        time.sleep(0.5)
        error = ser.read_all().decode().strip()
        
        if not error or error.startswith("0,"):  # No error
            break
            
        print(f"Error detected: {error}")
        if "USB-CDC" in error:
            print("USB-CDC interface error detected - waiting for interface to stabilize...")
            time.sleep(5)  # Longer delay for interface stability
        
        time.sleep(0.5)  # Delay between error checks
        return error  # Return error message for checking
    return None

def parse_three_phase_voltage(response):
    """Parse three phase voltage response"""
    if not response:
        return None, None, None
    try:
        # Split into three values (e.g., "0.1860.1640.175" -> ["0.186", "0.164", "0.175"])
        values = []
        start = 0
        while start < len(response):
            # Find next decimal point
            dot_pos = response.find('.', start)
            if dot_pos == -1:
                break
            # Take 4 chars after decimal (number + 3 decimal places)
            value = response[start:dot_pos+4]
            if value:
                values.append(float(value))
            start = dot_pos + 4
        
        if len(values) == 3:
            return values[0], values[1], values[2]
    except:
        pass
    return None, None, None

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
                
                if response:
                    return response
                elif attempt < max_retries - 1:
                    print(f"Attempt {attempt + 1}: No response, retrying...")
                    time.sleep(1)
                    continue
            else:
                # Check for errors after each command
                error = check_errors(ser)
                if error:
                    if attempt < max_retries - 1:
                        print(f"Attempt {attempt + 1}: Retrying due to error...")
                        time.sleep(1)
                        continue
                return ""
            
        except Exception as e:
            print(f"Command error: {str(e)}")
            if attempt < max_retries - 1:
                print(f"Attempt {attempt + 1}: Retrying...")
                time.sleep(1)
                continue
            raise
    
    return ""

def setup_agx(ser):
    """Basic AGX setup"""
    print("\nInitializing AGX device...")
    
    # Initial delay to ensure device is ready after restart
    print("Waiting for device to initialize...")
    time.sleep(10)  # Longer initial delay after reboot
    
    # Reset device and wait for it to complete
    send_command(ser, "*RST")
    time.sleep(2)
    
    # Clear status
    send_command(ser, "*CLS")
    time.sleep(1)
    
    # Set AC mode first
    send_command(ser, "SOURce:MODE AC")
    time.sleep(1)
    
    # Query mode to verify
    mode = send_command(ser, "SOURce:MODE?")
    if mode and "AC" in mode.upper():
        print("AC mode confirmed")
    else:
        print("Warning: Could not confirm AC mode")
        check_errors(ser)
        return False
    
    # Set frequency to 50Hz
    send_command(ser, "SOURce:FREQuency 50")
    time.sleep(1)
    
    # Enable voltage display
    send_command(ser, "DISPlay:VOLTage ON")
    time.sleep(1)
    
    return True

def set_three_phase_ac_voltage(ser, voltage):
    """Set three phase AC voltage"""
    try:
        print(f"\nSetting three phase AC voltage to {voltage}V...")
        
        # Disable output first
        send_command(ser, "OUTPut:STATe OFF")
        time.sleep(1)
        
        # Verify we're in AC mode
        mode = send_command(ser, "SOURce:MODE?")
        if not mode or "AC" not in mode.upper():
            print("Warning: Not in AC mode, setting AC mode...")
            send_command(ser, "SOURce:MODE AC")
            time.sleep(1)
        
        # Set voltage to 0 first for all phases
        for phase in range(1, 4):
            send_command(ser, f"SOURce:VOLTage:AC 0,(@{phase})")
            time.sleep(1)
        
        # Set target voltage for all phases
        for phase in range(1, 4):
            send_command(ser, f"SOURce:VOLTage:AC {voltage},(@{phase})")
            time.sleep(1)
        
        # Enable output
        send_command(ser, "OUTPut:STATe ON")
        time.sleep(2)
        
        # Verify output state
        output_state = send_command(ser, "OUTPut:STATe?")
        if output_state and output_state.strip() != "1":
            print("Warning: Output not enabled")
            check_errors(ser)
        
        # Measure actual voltage for all phases
        actual = send_command(ser, "MEASure:VOLTage:AC?")
        if actual:
            v1, v2, v3 = parse_three_phase_voltage(actual)
            if v1 is not None and v2 is not None and v3 is not None:
                print(f"Target voltage: {voltage}V AC")
                print(f"Measured voltage - Phase 1: {v1}V AC")
                print(f"Measured voltage - Phase 2: {v2}V AC")
                print(f"Measured voltage - Phase 3: {v3}V AC")
                if any(abs(v - voltage) > voltage * 0.1 for v in [v1, v2, v3]):  # 10% tolerance
                    print(f"Warning: Large voltage discrepancy detected")
                    check_errors(ser)
            else:
                print("Warning: Could not parse voltage measurements")
        else:
            print("Warning: No voltage measurements received")
            check_errors(ser)
    
    finally:
        # Ensure output is disabled before exiting
        print("\nDisabling output...")
        for phase in range(1, 4):
            send_command(ser, f"SOURce:VOLTage:AC 0,(@{phase})")
            time.sleep(1)
        send_command(ser, "OUTPut:STATe OFF")
        time.sleep(1)

def main():
    ser = None
    try:
        # Connect to AGX with longer timeout
        ser = serial.Serial(
            port='COM3',
            baudrate=9600,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=5
        )
        print("Connected to COM3")
        
        # Setup device
        if not setup_agx(ser):
            print("Failed to initialize device in AC mode")
            return
        
        # Test sequence - AC voltage
        print("\nStarting three phase AC voltage test sequence...")
        voltages = list(range(10, 101, 10))  # 10V to 100V in steps of 10V
        
        for voltage in voltages:
            set_three_phase_ac_voltage(ser, voltage)
            time.sleep(3)
        
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        if ser:
            try:
                # Disable output and set voltage to 0
                send_command(ser, "OUTPut:STATe OFF")
                for phase in range(1, 4):
                    send_command(ser, f"SOURce:VOLTage:AC 0,(@{phase})")
                ser.close()
                print("Serial connection closed")
            except:
                print("Error while closing connection")

if __name__ == "__main__":
    main()

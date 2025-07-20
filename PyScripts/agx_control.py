import serial
import time

def check_errors(ser):
    """Read and display all errors in the queue with enhanced USB error handling"""
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # Clear buffers before checking errors
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            
            ser.write("SYST:ERR?\n".encode())
            time.sleep(1)  # Longer wait for response
            error = ser.read_all().decode().strip()
            
            if not error or error.startswith("0,"):  # No error
                return None
                
            print(f"Error detected: {error}")
            
            if "USB-CDC" in error:
                print("USB-CDC interface error detected - attempting recovery...")
                # More aggressive USB error recovery
                ser.close()
                time.sleep(10)  # Longer delay for USB reset
                try:
                    ser.open()
                    # Re-initialize basic settings after USB recovery
                    send_command(ser, "*RST")
                    send_command(ser, "*CLS")
                    send_command(ser, "OUTP,OFF")  # Safety measure
                except Exception as e:
                    print(f"USB recovery failed: {str(e)}")
                    
            elif "timeout" in error.lower():
                print("Communication timeout - waiting before retry...")
                time.sleep(5)
            
            retry_count += 1
            if retry_count < max_retries:
                print(f"Retrying error check (attempt {retry_count + 1}/{max_retries})...")
                time.sleep(2)
            else:
                print("Max retries reached - device may need manual intervention")
                return error
                
        except Exception as e:
            print(f"Error checking system errors: {str(e)}")
            retry_count += 1
            if retry_count < max_retries:
                time.sleep(2)
            else:
                return f"Error check failed: {str(e)}"
    
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
    """Basic AGX setup using 3150Afx compatible commands"""
    print("\nInitializing AGX device...")
    
    # Initial delay to ensure device is ready after restart
    print("Waiting for device to initialize...")
    time.sleep(10)  # Longer initial delay after reboot
    
    # Reset and clear
    send_command(ser, "*RST")
    time.sleep(2)
    send_command(ser, "*CLS")
    time.sleep(1)
    
    # Basic initialization sequence based on 3150Afx
    commands = [
        "OUTPUT:AUTO,OFF",
        "INIT,OFF",
        "*LLO",
        "OUTP,OFF",
        "PROG:EXEC,0",
        "VOLT:MODE,AC",
        "FORM,3",  # Three phase mode
        "VOLT:ALC,ON",
        "CURR:LIM,10",
        "WAVEFORM,1",
        "SENS:PATH,0",
        "FREQ:LIM:MIN,45",
        "FREQ:LIM:MAX,1500",
        "FREQ,50",  # Set to 50Hz
        "VOLT:AC:LIM:MIN,0",
        "VOLT:AC:LIM:MAX,120",
        "COUPL,DIRECT",
        "PHAS2,120",
        "PHAS3,240",
        "RANG,1",
        "RAMP,0"
    ]
    
    for cmd in commands:
        send_command(ser, cmd)
        time.sleep(0.5)
        error = check_errors(ser)
        if error:
            print(f"Error during setup with command {cmd}")
            return False
    
    return True

def set_three_phase_ac_voltage(ser, voltage):
    """Set three phase AC voltage using 3150Afx compatible commands"""
    try:
        print(f"\nSetting three phase AC voltage to {voltage}V...")
        
        # Disable output first
        send_command(ser, "OUTP,OFF")
        time.sleep(1)
        
        # Set AC voltage for all phases using compatible command format
        send_command(ser, f"VOLT:AC,{voltage}")
        time.sleep(1)
        
        # Enable output
        send_command(ser, "OUTP,ON")
        time.sleep(2)
        
        # Verify output state with retry
        max_verify_attempts = 3
        output_enabled = False
        
        for attempt in range(max_verify_attempts):
            output_state = send_command(ser, "OUTP?")
            if output_state and output_state.strip() == "1":
                output_enabled = True
                break
            print(f"Output not enabled (attempt {attempt + 1}/{max_verify_attempts})")
            time.sleep(2)
            send_command(ser, "OUTP,ON")  # Retry enabling output
        
        if not output_enabled:
            print("Failed to enable output after multiple attempts")
            return
        
        # Wait for voltage to stabilize and hold display
        print("Holding display for 20 seconds...")
        time.sleep(20)  # Extended delay to hold display
        
        # Measure actual voltage using 3150Afx compatible commands
        actual = send_command(ser, "MEAS:VOLT?")
        if actual:
            try:
                # Parse voltage response (format may vary by device)
                voltages = [float(v) for v in actual.split(',') if v.strip()]
                if len(voltages) >= 3:
                    v1, v2, v3 = voltages[:3]
                    print(f"Target voltage: {voltage}V AC")
                    print(f"Measured voltage - Phase 1: {v1:.3f}V AC")
                    print(f"Measured voltage - Phase 2: {v2:.3f}V AC")
                    print(f"Measured voltage - Phase 3: {v3:.3f}V AC")
                    
                    # Check if voltages are within expected range
                    tolerance = 0.1  # 10% tolerance
                    if any(abs(v - voltage) > voltage * tolerance for v in [v1, v2, v3]):
                        print("Warning: Voltage outside expected range")
                        # Try to adjust voltage if needed
                        send_command(ser, f"VOLT,{voltage}")
                        time.sleep(1)
                else:
                    print("Warning: Incomplete voltage measurements received")
            except ValueError as e:
                print(f"Warning: Could not parse voltage measurements: {str(e)}")
        else:
            print("Warning: No voltage measurements received")
    
    finally:
        # Safe shutdown sequence for voltage setting
        print("\nPerforming safe shutdown...")
        try:
            # Set voltage to 0 first
            send_command(ser, "VOLT,0")
            time.sleep(1)
            
            # Disable output
            send_command(ser, "OUTP,OFF")
            time.sleep(1)
            
            # Verify output is off
            output_state = send_command(ser, "OUTP?")
            if output_state and output_state.strip() != "0":
                print("Warning: Output may still be enabled")
                # Force disable
                send_command(ser, "OUTP,OFF")
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")

def setup_agx_dc(ser):
    """DC mode setup using 3150Afx compatible commands"""
    print("\nInitializing AGX device for DC mode...")
    
    # Reset and clear
    send_command(ser, "*RST")
    time.sleep(2)
    send_command(ser, "*CLS")
    time.sleep(1)
    
    # Initialization sequence from 3150Afx ThreePhaseControlsDC
    commands = [
        "OUTP,OFF",
        "INIT,OFF",
        "*LLO",
        "PROG:EXEC,0",
        "VOLT:MODE,DC",  # Set DC mode
        "FORM,3",        # Three phase mode
        "VOLT:ALC,ON",   # Automatic level control
        "CURR:LIM,10",   # Current limit
        "WAVEFORM,1",
        "SENS:PATH,0",
        "VOLT:DC:LIM:MIN,0",   # DC voltage limits
        "VOLT:DC:LIM:MAX,120", # Max 120V DC
        "COUPL,DIRECT",
        "PHAS2,120",     # Phase angles
        "PHAS3,240",
        "RANG,1",
        "RAMP,0"
        # Note: OUTP,ON will be set after verification
    ]
    
    for cmd in commands:
        send_command(ser, cmd)
        time.sleep(0.5)
        error = check_errors(ser)
        if error:
            print(f"Error during DC setup with command {cmd}")
            return False
    
    return True

def set_three_phase_dc_voltage(ser, voltage):
    """Set three phase DC voltage using 3150Afx compatible commands"""
    try:
        # Validate voltage range for DC mode
        if not 0 <= voltage <= 120:
            print(f"Warning: Voltage {voltage}V outside valid DC range (0-120V)")
            voltage = min(max(voltage, 0), 120)
            print(f"Adjusting to {voltage}V")
        
        print(f"\nSetting three phase DC voltage to {voltage}V...")
        
        # Disable output first
        send_command(ser, "OUTP,OFF")
        time.sleep(1)
        
        # Verify DC mode is active
        mode = send_command(ser, "VOLT:MODE?")
        if mode and "DC" not in mode.upper():
            print("Warning: Device not in DC mode, switching to DC mode...")
            send_command(ser, "VOLT:MODE,DC")
            time.sleep(1)
        
        # Set voltage for all phases
        send_command(ser, f"VOLT,{voltage}")
        time.sleep(1)
        
        # Enable output
        send_command(ser, "OUTP,ON")
        time.sleep(2)
        
        # Verify output state with retry
        max_verify_attempts = 3
        output_enabled = False
        
        for attempt in range(max_verify_attempts):
            output_state = send_command(ser, "OUTP?")
            if output_state and output_state.strip() == "1":
                output_enabled = True
                break
            print(f"Output not enabled (attempt {attempt + 1}/{max_verify_attempts})")
            time.sleep(2)
            send_command(ser, "OUTP,ON")
        
        if not output_enabled:
            print("Failed to enable output after multiple attempts")
            return
        
        # Wait for voltage to stabilize
        print("Waiting for voltage to stabilize...")
        time.sleep(5)
        
        # Take multiple measurements for accuracy (like in C# implementation)
        measurements = {
            'DC1': [],
            'DC2': [],
            'DC3': []
        }
        
        print("Taking measurements...")
        for _ in range(10):
            for phase in ['DC1', 'DC2', 'DC3']:
                response = send_command(ser, f"MEAS:VOLT:{phase}?")
                if response:
                    try:
                        measurements[phase].append(float(response.strip()))
                    except ValueError:
                        print(f"Warning: Invalid measurement for phase {phase}")
        
        # Calculate averages
        averages = {}
        for phase, values in measurements.items():
            if values:
                averages[phase] = sum(values) / len(values)
                print(f"Measured voltage - Phase {phase[-1]}: {averages[phase]:.3f}V DC")
            else:
                print(f"Warning: No valid measurements for phase {phase}")
        
        # Check if voltages are within expected range
        tolerance = 0.1  # 10% tolerance
        for phase, avg in averages.items():
            if abs(avg - voltage) > voltage * tolerance:
                print(f"Warning: Phase {phase[-1]} voltage outside expected range")
                # Try to adjust voltage if needed
                send_command(ser, f"VOLT,{voltage}")
                time.sleep(1)
    
    finally:
        # Safe shutdown sequence for voltage setting
        print("\nPerforming safe shutdown...")
        try:
            # Set voltage to 0 first
            send_command(ser, "VOLT,0")
            time.sleep(1)
            
            # Disable output
            send_command(ser, "OUTP,OFF")
            time.sleep(1)
            
            # Verify output is off
            output_state = send_command(ser, "OUTP?")
            if output_state and output_state.strip() != "0":
                print("Warning: Output may still be enabled")
                # Force disable
                send_command(ser, "OUTP,OFF")
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")

def main():
    ser = None
    try:
        # Connect to AGX with longer timeout and higher baud rate
        ser = serial.Serial(
            port='COM3',
            baudrate=115200,  # Increased from 9600 for faster communication
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=5
        )
        print("Connected to COM3")
        
        # AC Test Sequence
        print("\n=== Starting AC Test Sequence ===")
        if not setup_agx(ser):
            print("Failed to initialize device in AC mode")
            return
        
        print("\nStarting three phase AC voltage test sequence...")
        voltages = list(range(10, 121, 10))  # 10V to 120V in steps of 10V
        
        # First AC test with 30 second settling time
        print(f"\nSetting and holding AC voltage at {voltages[0]}V for 30 seconds...")
        set_three_phase_ac_voltage(ser, voltages[0])
        time.sleep(30)  # Initial longer settling time
        
        # Remaining AC tests with 15 second settling time
        for voltage in voltages[1:]:
            print(f"\nSetting and holding AC voltage at {voltage}V for 15 seconds...")
            set_three_phase_ac_voltage(ser, voltage)
            time.sleep(15)
            
        # DC Test Sequence
        print("\n=== Starting DC Test Sequence ===")
        if not setup_agx_dc(ser):
            print("Failed to initialize device in DC mode")
            return
            
        print("\nStarting three phase DC voltage test sequence...")
        voltages = list(range(10, 121, 10))  # 10V to 120V in steps of 10V
        
        # First DC test with 30 second settling time
        print(f"\nSetting and holding DC voltage at {voltages[0]}V for 30 seconds...")
        set_three_phase_dc_voltage(ser, voltages[0])
        time.sleep(30)  # Initial longer settling time
        
        # Remaining DC tests with 15 second settling time
        for voltage in voltages[1:]:
            print(f"\nSetting and holding DC voltage at {voltage}V for 15 seconds...")
            set_three_phase_dc_voltage(ser, voltage)
            time.sleep(15)
        
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        if ser:
            try:
                # Complete shutdown sequence
                print("\nPerforming final shutdown...")
                
                # Set voltage to 0
                send_command(ser, "VOLT,0")
                time.sleep(1)
                
                # Disable output
                send_command(ser, "OUTP,OFF")
                time.sleep(1)
                
                # Reset device to safe state
                send_command(ser, "*RST")
                time.sleep(1)
                
                # Clear status
                send_command(ser, "*CLS")
                
                # Close connection
                ser.close()
                print("Device safely shut down and connection closed")
            except Exception as e:
                print(f"Error during final shutdown: {str(e)}")
                try:
                    ser.close()
                except:
                    pass

if __name__ == "__main__":
    main()

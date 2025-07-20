"""
Simple test script to verify AGX and N4L communications
"""
import serial
import time

def test_basic_communication():
    """Run basic communication tests with both devices"""
    try:
        print("\nInitializing devices...")
        # Setup AGX communication
        agx = serial.Serial(
            port='COM3',
            baudrate=115200,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=1
        )
        
        if not agx.is_open:
            print("Failed to open COM3")
            return
            
        # Test AGX communication
        print("\nTesting AGX communication:")
        try:
            def send_command(cmd):
                print(f"Sending: {cmd}")
                agx.write(f"{cmd}\n".encode())
                time.sleep(0.5)
                response = agx.read_all().decode().strip()
                print(f"Response: {response}")
                return response

            # Query device ID
            print("\nTesting AGX communication:")
            send_command("*IDN?")
            
            # Set a simple voltage and measure it
            print("\nTesting voltage setting and measurement:")
            # Start with output off
            send_command(":OUTP,OFF")
            
            # Configure for AC mode
            send_command(":VOLT:MODE,AC")
            send_command(":FORM,1")  # Single phase
            send_command(":FREQ,50")  # 50Hz
            send_command(":VOLT:AC:LIM:MIN,0")
            send_command(":VOLT:AC:LIM:MAX,300")
            
            # Set 10V output
            print("Setting 10V AC output...")
            send_command(":VOLT:AC,10")
            send_command(":OUTP,ON")
            time.sleep(5)  # Wait for stabilization
            
            # Measure voltage
            voltage = send_command(":MEAS:VOLT:AC1?")
            print(f"Measured voltage: {voltage}V")
            
            # Turn off output
            send_command(":OUTP,OFF")
            
        except Exception as e:
            print(f"Error testing AGX: {e}")
            
        # Test N4L communication
        print("\nTesting N4L communication:")
        try:
            # Setup N4L communication
            n4l = serial.Serial(
                port='COM4',  # Adjust if needed
                baudrate=115200,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            
            def send_n4l_command(cmd):
                print(f"Sending to N4L: {cmd}")
                n4l.write(f"{cmd}\n".encode())
                time.sleep(0.5)
                response = n4l.read_all().decode().strip()
                print(f"N4L Response: {response}")
                return response

            # Reset and configure
            send_n4l_command("*RST")
            
            # Query ID
            send_n4l_command("*IDN?")
            
            # Test basic measurement setup
            print("\nConfiguring N4L measurements...")
            commands = [
                "SYST;POWER",               # Power meter mode
                "COUPLI,PHASE1,AC+DC",      # AC+DC coupling
                "BANDWI,WIDE",              # Wide bandwidth
                "RESOLU,HIGH",              # High resolution
                "SPEED,HIGH",               # High speed
                "DISPLAY,VOLTAGE",          # Voltage display mode
            ]
            
            for cmd in commands:
                send_n4l_command(cmd)
                
            # Take a voltage reading
            print("\nTaking voltage measurement...")
            send_n4l_command("MULTIL,1,1,50")  # Configure for RMS voltage
            response = send_n4l_command("MULTIL?")
            print(f"N4L measurement: {response}")
            
            n4l.close()
            
        except Exception as e:
            print(f"Error testing N4L: {e}")
            
    finally:
        print("\nCleaning up...")
        if 'agx' in locals() and agx.is_open:
            agx.close()
        if 'n4l' in locals() and n4l.is_open:
            n4l.close()
        print("Test complete")

if __name__ == "__main__":
    test_basic_communication()

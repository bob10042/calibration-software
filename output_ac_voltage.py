import pyvisa
import time
import sys

class ACPowerSupply:
    def __init__(self):
        try:
            self.rm = pyvisa.ResourceManager()
            resources = self.rm.list_resources()
            gpib_devices = [res for res in resources if 'GPIB' in res]
            
            if not gpib_devices:
                raise Exception("No GPIB devices found")
                
            self.instrument = self.rm.open_resource(gpib_devices[0])
            self.instrument.timeout = 5000
            print(f"Connected to: {self.instrument.query('*IDN?')}")
            
        except Exception as e:
            print(f"Initialization error: {str(e)}")
            sys.exit(1)
            
    def setup_ac_mode(self):
        """Configure the power supply for AC output"""
        commands = [
            ":VOLT:MODE,AC",    # Set AC mode
            ":FORM,1",          # Single phase mode
            ":VOLT:ALC,ON",     # Enable automatic level control
            ":CURR:LIM,10",     # Set current limit to 10A
            ":VOLT:AC:LIM:MIN,0",
            ":VOLT:AC:LIM:MAX,300",  # Set max voltage to 300V
            ":FREQ,50"          # Set to 50Hz
        ]
        
        for cmd in commands:
            self.instrument.write(cmd)
            time.sleep(0.1)
    
    def set_frequency(self, freq):
        """Set the output frequency"""
        if not 45 <= freq <= 1500:
            raise ValueError("Frequency must be between 45 and 1500 Hz")
        self.instrument.write(f":FREQ,{freq}")
        time.sleep(0.1)
    
    def set_voltage(self, voltage):
        """Set the AC voltage"""
        if not 0 <= voltage <= 300:
            raise ValueError("Voltage must be between 0 and 300V")
        self.instrument.write(f":VOLT:AC,{voltage}")
        time.sleep(0.1)
    
    def enable_output(self, enable=True):
        """Enable or disable the output"""
        self.instrument.write(f":OUTP,{'ON' if enable else 'OFF'}")
        time.sleep(0.1)
    
    def measure_voltage(self):
        """Measure the actual output voltage"""
        return float(self.instrument.query(":MEAS:VOLT:AC1?"))
    
    def shutdown(self):
        """Safely shutdown the power supply"""
        try:
            self.set_voltage(0)
            self.enable_output(False)
            self.instrument.close()
            self.rm.close()
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")

def main():
    ps = None
    try:
        ps = ACPowerSupply()
        
        # Setup AC mode
        print("Setting up AC mode...")
        ps.setup_ac_mode()
        
        # Set initial voltage to 0
        ps.set_voltage(0)
        
        # Enable output
        print("Enabling output...")
        ps.enable_output(True)
        
        while True:
            try:
                voltage = float(input("\nEnter desired voltage (0-300V) or -1 to exit: "))
                if voltage == -1:
                    break
                    
                if 0 <= voltage <= 300:
                    print(f"\nSetting voltage to {voltage}V...")
                    ps.set_voltage(voltage)
                    time.sleep(1)  # Wait for voltage to stabilize
                    
                    # Measure and display actual voltage
                    measured = ps.measure_voltage()
                    print(f"Measured voltage: {measured:.2f}V")
                else:
                    print("Voltage must be between 0 and 300V")
                    
            except ValueError:
                print("Please enter a valid number")
                
    except KeyboardInterrupt:
        print("\nProgram interrupted by user")
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        if ps:
            print("\nShutting down...")
            ps.shutdown()
            print("Shutdown complete")

if __name__ == "__main__":
    main()

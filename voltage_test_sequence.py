import pyvisa
import time
import sys
from datetime import datetime

class PowerSupplyTester:
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
            
    def setup_dc_mode(self):
        """Configure the power supply for DC output"""
        commands = [
            ":VOLT:MODE,DC",    # Set DC mode
            ":FORM,1",          # Single phase mode
            ":VOLT:ALC,ON",     # Enable automatic level control
            ":CURR:LIM,10",     # Set current limit to 10A
            ":VOLT:DC:LIM:MIN,0",
            ":VOLT:DC:LIM:MAX,300"  # Set max voltage to 300V
        ]
        
        for cmd in commands:
            self.instrument.write(cmd)
            time.sleep(0.1)
    
    def set_voltage(self, voltage, mode='AC'):
        """Set the voltage"""
        if not 0 <= voltage <= 300:
            raise ValueError("Voltage must be between 0 and 300V")
        self.instrument.write(f":VOLT:{mode},{voltage}")
        time.sleep(0.1)
    
    def enable_output(self, enable=True):
        """Enable or disable the output"""
        self.instrument.write(f":OUTP,{'ON' if enable else 'OFF'}")
        time.sleep(0.1)
    
    def measure_voltage(self, mode='AC'):
        """Measure the actual output voltage"""
        command = f":MEAS:VOLT:{mode}1?"
        return float(self.instrument.query(command))
    
    def run_voltage_test(self, mode='AC'):
        """Run voltage test sequence"""
        print(f"\nStarting {mode} voltage test sequence")
        print("=====================================")
        
        # Setup appropriate mode
        if mode == 'AC':
            self.setup_ac_mode()
        else:
            self.setup_dc_mode()
            
        # Set initial voltage to 0 and enable output
        self.set_voltage(0, mode)
        self.enable_output(True)
        
        # Test voltages from 10V to 120V in 10V steps
        for voltage in range(10, 121, 10):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"\n{timestamp}")
            print(f"Setting {mode} voltage to {voltage}V")
            
            # Set the voltage
            self.set_voltage(voltage, mode)
            
            # Take readings during 15 second stabilization period
            print("Starting 15 second stabilization period with readings:")
            readings = []
            for i in range(5):  # Take 5 readings over 15 seconds
                time.sleep(3)  # Wait 3 seconds between readings
                measured = self.measure_voltage(mode)
                readings.append(measured)
                print(f"Reading {i+1} at {i*3+3}s: {measured:.3f}V")
            
            # Calculate statistics
            avg_voltage = sum(readings) / len(readings)
            max_dev = max(abs(v - voltage) for v in readings)
            print(f"\nFinal Statistics:")
            print(f"Average {mode} voltage: {avg_voltage:.3f}V")
            print(f"Maximum deviation: {max_dev:.3f}V")
            print(f"Target voltage: {voltage:.1f}V")
            
            # Wait remaining 5 seconds before next voltage
            print("Waiting 5 more seconds before next voltage...")
            time.sleep(5)
        
        # Set voltage back to 0
        self.set_voltage(0, mode)
    
    def shutdown(self):
        """Safely shutdown the power supply"""
        try:
            self.set_voltage(0, 'AC')
            self.set_voltage(0, 'DC')
            self.enable_output(False)
            self.instrument.close()
            self.rm.close()
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")

def main():
    ps = None
    try:
        ps = PowerSupplyTester()
        
        # Run AC voltage tests
        ps.run_voltage_test(mode='AC')
        
        # Short pause between AC and DC tests
        print("\nPausing for 5 seconds before DC tests...")
        time.sleep(5)
        
        # Run DC voltage tests
        ps.run_voltage_test(mode='DC')
        
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

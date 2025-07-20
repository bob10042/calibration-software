import pyvisa
import pandas as pd
import time
from datetime import datetime
import sys
import csv

class UKASTestRunner:
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
            ":FORM,3",          # Three phase mode
            ":VOLT:ALC,ON",     # Enable automatic level control
            ":CURR:LIM,10",     # Set current limit to 10A
            ":VOLT:AC:LIM:MIN,0",
            ":VOLT:AC:LIM:MAX,1200",  # Set max voltage to 1200V for AC
            ":FREQ,50"          # Set to 50Hz
        ]
        
        for cmd in commands:
            self.instrument.write(cmd)
            time.sleep(0.1)
    
    def setup_dc_mode(self):
        """Configure the power supply for DC output"""
        commands = [
            ":VOLT:MODE,DC",    # Set DC mode
            ":FORM,3",          # Three phase mode
            ":VOLT:ALC,ON",     # Enable automatic level control
            ":CURR:LIM,10",     # Set current limit to 10A
            ":VOLT:DC:LIM:MIN,0",
            ":VOLT:DC:LIM:MAX,425"  # Set max voltage to 425V for DC
        ]
        
        for cmd in commands:
            self.instrument.write(cmd)
            time.sleep(0.1)
    
    def display_setup_instructions(self, mode, phase_config):
        """Display setup instructions for the current test group"""
        print("\n" + "="*50)
        print(f"Test Setup Instructions for {mode} Mode, {phase_config}")
        print("="*50)
        print("1. Ensure proper safety measures are in place")
        print("2. Verify all connections:")
        if mode == "AC":
            print("   - Connect voltage measurement device for AC measurements")
            print("   - Set measurement device to AC mode")
            print("   - Verify 50Hz frequency setting")
        else:
            print("   - Connect voltage measurement device for DC measurements")
            print("   - Set measurement device to DC mode")
        print("3. Check all cable connections are secure")
        print("4. Ensure proper grounding")
        print("\nPress Enter when ready to proceed...")
        input()
    
    def set_voltage(self, voltage, mode='AC'):
        """Set the voltage"""
        if mode == 'AC':
            self.instrument.write(f":VOLT:AC,{voltage}")
        else:
            self.instrument.write(f":VOLT:DC,{voltage}")
        time.sleep(0.1)
    
    def measure_voltage(self, mode='AC', phase=1):
        """Measure the voltage on specified phase"""
        command = f":MEAS:VOLT:{mode}{phase}?"
        return float(self.instrument.query(command))
    
    def run_test_sequence(self, csv_file):
        """Run the complete test sequence from CSV file"""
        try:
            # Read test sequence
            df = pd.read_csv(csv_file, comment='#')
            
            # Create results file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            results_file = f"test_results_{timestamp}.csv"
            
            with open(results_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Test Point', 'Mode', 'Phase', 'Target Voltage', 'Measured Voltage', 'Timestamp'])
            
            # Group tests by mode
            current_mode = None
            current_phase_config = None
            
            for index, test in df.iterrows():
                # Check if mode or phase config changed
                if test['mode'] != current_mode or test['phase_config'] != current_phase_config:
                    current_mode = test['mode']
                    current_phase_config = test['phase_config']
                    
                    # Setup for new mode
                    if current_mode == 'AC':
                        self.setup_ac_mode()
                    else:
                        self.setup_dc_mode()
                    
                    # Display setup instructions
                    self.display_setup_instructions(current_mode, current_phase_config)
                
                # Extract phase number from test point (A-N, B-N, C-N)
                phase = 1  # Default to phase 1
                if "B-N" in test['test_point']:
                    phase = 2
                elif "C-N" in test['test_point']:
                    phase = 3
                
                print(f"\nExecuting: {test['test_point']}")
                print(f"Setting {test['mode']} voltage to {test['voltage']}V")
                
                # Set voltage
                self.set_voltage(test['voltage'], test['mode'])
                
                # Wait 20 seconds
                print("Waiting 20 seconds for stabilization...")
                time.sleep(20)
                
                # Take measurement
                measured = self.measure_voltage(test['mode'], phase)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                print(f"Measured voltage: {measured:.3f}V")
                
                # Save results
                with open(results_file, 'a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        test['test_point'],
                        test['mode'],
                        phase,
                        test['voltage'],
                        measured,
                        timestamp
                    ])
                
                # Set voltage back to 0 between tests
                self.set_voltage(0, test['mode'])
                time.sleep(2)
                
        except Exception as e:
            print(f"Error during test sequence: {str(e)}")
            self.shutdown()
            sys.exit(1)
    
    def shutdown(self):
        """Safely shutdown the power supply"""
        try:
            self.set_voltage(0, 'AC')
            self.set_voltage(0, 'DC')
            self.instrument.write(":OUTP,OFF")
            self.instrument.close()
            self.rm.close()
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")

def main():
    runner = None
    try:
        runner = UKASTestRunner()
        
        print("\nUKAS Voltage Test Sequence")
        print("=" * 50)
        print("This script will run through the UKAS voltage test sequence.")
        print("Please ensure all safety measures are in place before proceeding.")
        print("\nPress Enter to begin...")
        input()
        
        runner.run_test_sequence('ukas_voltage_tests_new.csv')
        
    except KeyboardInterrupt:
        print("\nTest sequence interrupted by user")
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        if runner:
            print("\nShutting down...")
            runner.shutdown()
            print("Shutdown complete")

if __name__ == "__main__":
    main()

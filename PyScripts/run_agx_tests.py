"""
AGX Test Runner using configurations from agx_test_configs.py
This script implements the test procedures for different voltage and current modes
"""

import time
from typing import List, Dict, Any
import pyvisa
from agx_test_configs import AGXConfigurations

class AGXTestRunner:
    def __init__(self):
        self.rm = None
        self.agx = None
        self.n4l = None  # Newton's 4th Power Analyzer
        self.configs = AGXConfigurations()
        self.baud_rate = 115200  # Increased from 9600 for faster communication
        
    def setup_instruments(self, gpib_address: int = 1):
        """Initialize and setup communication with instruments"""
        try:
            # Setup VISA communication with AGX
            self.rm = pyvisa.ResourceManager()
            resource = f'GPIB0::{gpib_address}::INSTR'
            self.agx = self.rm.open_resource(resource)
            
            # Basic instrument setup
            self.agx.write('*RST')  # Reset
            self.agx.write('*CLS')  # Clear status
            
            # Setup Newton's 4th Power Analyzer
            self._setup_newton_4th()
            
            return True
        except Exception as e:
            print(f"Error setting up instruments: {e}")
            return False
            
    def _setup_newton_4th(self):
        """Configure Newton's 4th Power Analyzer with higher baud rate for faster communication"""
        try:
            # Configure and open serial port
            import serial
            self.n4l = serial.Serial(
                port=None,  # Port name will be set later
                baudrate=self.baud_rate,  # Using 115200 instead of 9600
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=2,
                write_timeout=2,
                xonxoff=False,
                rtscts=True,
                dsrdtr=True
            )
            
            # Get available ports
            from serial.tools import list_ports
            ports = list(list_ports.comports())
            for port in ports:
                if "N4L" in port.description or "Newton" in port.description:
                    self.n4l.port = port.device
                    break
            
            if not self.n4l.port:
                raise Exception("Newton's 4th Power Analyzer not found")
                
            # Open port
            self.n4l.open()
            if not self.n4l.is_open:
                raise Exception("Failed to open serial port")
                
            print(f"Connected to N4L at {self.n4l.port} with {self.baud_rate} baud")
            
            # Initialize device
            for cmd in self.configs.NEWTON_4TH_INIT['commands']:
                self.n4l.write(f"{cmd}\n".encode())
                time.sleep(0.2)  # Small delay between commands
                
            # Setup multilogs
            for cmd in self.configs.NEWTON_4TH_INIT['multilog_setup']:
                self.n4l.write(f"{cmd}\n".encode())
                time.sleep(0.2)
                
            # Allow time for settings to take effect
            time.sleep(5)
            
            return True
            
        except Exception as e:
            print(f"Error setting up Newton's 4th: {e}")
            return False
        
    def configure_mode(self, mode_config: Dict[str, Any]):
        """Configure AGX for specific mode"""
        try:
            for cmd in mode_config['commands']:
                self.agx.write(cmd)
                time.sleep(0.2)
            return True
        except Exception as e:
            print(f"Error configuring mode: {e}")
            return False
            
    def take_measurements(self, measurement_cmds: List[str], samples: int = 10) -> Dict[str, float]:
        """Take measurements with averaging"""
        results = {}
        for cmd in measurement_cmds:
            total = 0
            for _ in range(samples):
                try:
                    value = float(self.agx.query(cmd))
                    total += value
                    time.sleep(0.1)
                except Exception as e:
                    print(f"Error taking measurement {cmd}: {e}")
                    return None
            results[cmd] = total / samples
        return results
        
    def run_three_phase_ac_test(self, test_points: List[float]):
        """Run three phase AC voltage test"""
        print("\nRunning Three Phase AC Test")
        
        # Configure for three phase AC
        if not self.configure_mode(self.configs.THREE_PHASE_AC):
            return None
            
        # Initial stabilization
        time.sleep(self.configs.TEST_FLOWS['three_phase_ac']['stabilization_time'] / 1000)
        
        results = []
        for voltage in test_points:
            print(f"\nTesting at {voltage}V")
            
            # Set voltage
            self.agx.write(f":VOLT:AC,{voltage}")
            
            # Wait for stabilization
            time.sleep(self.configs.TEST_FLOWS['three_phase_ac']['measurement_delay'] / 1000)
            
            # Take measurements
            measurements = self.take_measurements([
                self.configs.MEASUREMENT_METHODS['voltage_ac1'],
                self.configs.MEASUREMENT_METHODS['voltage_ac2'],
                self.configs.MEASUREMENT_METHODS['voltage_ac3']
            ])
            
            if measurements:
                results.append({
                    'set_point': voltage,
                    'measurements': measurements
                })
                
        return results
        
    def run_three_phase_dc_test(self, test_points: List[float]):
        """Run three phase DC voltage test"""
        print("\nRunning Three Phase DC Test")
        
        # Configure for three phase DC
        if not self.configure_mode(self.configs.THREE_PHASE_DC):
            return None
            
        # Initial stabilization
        time.sleep(self.configs.TEST_FLOWS['three_phase_dc']['stabilization_time'] / 1000)
        
        results = []
        for voltage in test_points:
            print(f"\nTesting at {voltage}V")
            
            # Set voltage
            self.agx.write(f":VOLT:DC,{voltage}")
            
            # Wait for stabilization
            time.sleep(self.configs.TEST_FLOWS['three_phase_dc']['measurement_delay'] / 1000)
            
            # Take measurements
            measurements = self.take_measurements([
                self.configs.MEASUREMENT_METHODS['voltage_dc1'],
                self.configs.MEASUREMENT_METHODS['voltage_dc2'],
                self.configs.MEASUREMENT_METHODS['voltage_dc3']
            ])
            
            if measurements:
                results.append({
                    'set_point': voltage,
                    'measurements': measurements
                })
                
        return results
        
    def run_split_phase_test(self, test_points: List[float], mode: str = 'AC'):
        """Run split phase voltage test"""
        print(f"\nRunning Split Phase {mode} Test")
        
        # Configure for split phase
        config = self.configs.SPLIT_PHASE_AC if mode == 'AC' else self.configs.SPLIT_PHASE_DC
        if not self.configure_mode(config):
            return None
            
        # Initial stabilization
        flow_key = f'split_phase_{mode.lower()}'
        time.sleep(self.configs.TEST_FLOWS[flow_key]['stabilization_time'] / 1000)
        
        results = []
        for voltage in test_points:
            print(f"\nTesting at {voltage}V")
            
            # Set voltage
            self.agx.write(f":VOLT:{mode},{voltage}")
            
            # Wait for stabilization
            time.sleep(self.configs.TEST_FLOWS[flow_key]['measurement_delay'] / 1000)
            
            # Take measurements
            measurements = self.take_measurements([
                self.configs.MEASUREMENT_METHODS['voltage_line_line']
            ])
            
            if measurements:
                results.append({
                    'set_point': voltage,
                    'measurements': measurements
                })
                
        return results
        
    def run_single_phase_test(self, test_points: List[float], mode: str = 'AC'):
        """Run single phase voltage test"""
        print(f"\nRunning Single Phase {mode} Test")
        print("NOTE: Ensure three phase outputs are linked before proceeding")
        
        input("Press Enter to continue...")
        
        # Configure for single phase
        config = self.configs.SINGLE_PHASE_AC if mode == 'AC' else self.configs.SINGLE_PHASE_DC
        if not self.configure_mode(config):
            return None
            
        # Initial stabilization
        flow_key = f'single_phase_{mode.lower()}'
        time.sleep(self.configs.TEST_FLOWS[flow_key]['stabilization_time'] / 1000)
        
        results = []
        for voltage in test_points:
            print(f"\nTesting at {voltage}V")
            
            # Set voltage
            self.agx.write(f":VOLT:{mode},{voltage}")
            
            # Wait for stabilization
            time.sleep(self.configs.TEST_FLOWS[flow_key]['measurement_delay'] / 1000)
            
            # Take measurements
            meas_method = 'voltage_ac1' if mode == 'AC' else 'voltage_dc1'
            measurements = self.take_measurements([
                self.configs.MEASUREMENT_METHODS[meas_method]
            ])
            
            if measurements:
                results.append({
                    'set_point': voltage,
                    'measurements': measurements
                })
                
        return results
        
    def cleanup(self):
        """Clean up and close connections"""
        try:
            if self.agx:
                self.agx.write(':OUTP,OFF')  # Turn off output
                self.agx.write('*GTL')       # Go to local mode
                self.agx.write('*RST')       # Reset
                self.agx.close()
                
            if self.n4l:
                self.n4l.write('KEYBOARD,ENABLE')  # Re-enable keyboard
                self.n4l.write('*RST')            # Reset
                self.n4l.close()
                
            if self.rm:
                self.rm.close()
                
        except Exception as e:
            print(f"Error during cleanup: {e}")

def main():
    """Example usage of the AGX test runner"""
    # Test points from UKAS voltage tests
    ac_test_points = [10, 25, 50, 75, 100, 115, 135, 150, 200, 240, 270, 300]
    dc_test_points = [0, 25, 50, 75, 100, 120, 150, 200, 250, 300, 350, 400, 425]
    
    runner = AGXTestRunner()
    
    try:
        # Setup instruments
        if not runner.setup_instruments():
            print("Failed to setup instruments")
            return
            
        # Run three phase tests
        ac_results = runner.run_three_phase_ac_test(ac_test_points)
        dc_results = runner.run_three_phase_dc_test(dc_test_points)
        
        # Run split phase tests
        split_ac_results = runner.run_split_phase_test(ac_test_points, 'AC')
        split_dc_results = runner.run_split_phase_test(dc_test_points, 'DC')
        
        # Run single phase tests
        single_ac_results = runner.run_single_phase_test(ac_test_points, 'AC')
        single_dc_results = runner.run_single_phase_test(dc_test_points, 'DC')
        
        # Print results
        print("\nTest Results:")
        print("Three Phase AC:", ac_results)
        print("Three Phase DC:", dc_results)
        print("Split Phase AC:", split_ac_results)
        print("Split Phase DC:", split_dc_results)
        print("Single Phase AC:", single_ac_results)
        print("Single Phase DC:", single_dc_results)
        
    finally:
        runner.cleanup()

if __name__ == "__main__":
    main()

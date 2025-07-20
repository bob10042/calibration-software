import pyvisa
import time
from typing import List, Tuple, Dict
import csv
from datetime import datetime
import argparse
import statistics
import os

class AGXVoltageTest:
    def __init__(self):
        self.rm = None
        self.agx = None
        self.connected = False
        self.test_name = "voltage_test"

    def connect_to_agx(self, gpib_address: int = 1) -> bool:
        """Connect to AGX via GPIB"""
        print("\n=== Connecting to AGX ===")
        try:
            self.rm = pyvisa.ResourceManager()
            resource = f'GPIB0::{gpib_address}::INSTR'
            self.agx = self.rm.open_resource(resource)
            
            # Test connection with IDN query
            idn = self.agx.query('*IDN?')
            print(f"Connected to: {idn.strip()}")
            
            # Reset device and clear status
            self.agx.write('*RST')
            self.agx.write('*CLS')
            
            self.connected = True
            return True
            
        except Exception as e:
            print(f"Error connecting to AGX: {e}")
            return False

    def configure_mode(self, mode: str):
        """Configure AGX operating mode"""
        print(f"\n=== Configuring Mode: {mode} ===")
        try:
            mode = mode.upper()
            if mode == "AC":
                self.agx.write('CONF:MODE AC')
            elif mode == "DC":
                self.agx.write('CONF:MODE DC')
            elif mode == "AC+DC":
                self.agx.write('CONF:MODE ACDC')
            else:
                print(f"Unknown mode: {mode}")
                return False
            print(f"Mode set to {mode}")
            return True
        except Exception as e:
            print(f"Error setting mode: {e}")
            return False

    def setup_phase_config(self, phase_config: str, frequency: float = 50):
        """Configure phase configuration"""
        print(f"\n=== Setting up Phase Configuration: {phase_config} ===")
        try:
            if phase_config.upper() == "3-PHASE":
                self.agx.write('CONF:NOUT 3')
                self.agx.write(f'FREQ {frequency}')
            elif phase_config.upper() == "1-PHASE":
                self.agx.write('CONF:NOUT 1')
                self.agx.write(f'FREQ {frequency}')
            else:
                print(f"Unknown phase configuration: {phase_config}")
                return False
            
            # Set voltage unit to RMS
            self.agx.write('VOLT:UNIT VRMS')
            print(f"Phase configuration set to {phase_config}")
            return True
        except Exception as e:
            print(f"Error setting phase configuration: {e}")
            return False

    def set_voltage(self, voltage: float, phase: int = None):
        """Set voltage for specified phase or all phases"""
        try:
            if phase is None:
                # Set all phases to same voltage
                self.agx.write(f'VOLT {voltage}')
                print(f"Set all phases to {voltage}V")
            else:
                # Set specific phase voltage
                self.agx.write(f'VOLT{phase} {voltage}')
                print(f"Set phase {phase} to {voltage}V")
            return True
        except Exception as e:
            print(f"Error setting voltage: {e}")
            return False

    def measure_voltage(self, phase: int = None, samples: int = 1) -> float:
        """Measure voltage on specified phase or all phases with optional averaging"""
        try:
            if phase is None:
                # Measure all phases
                voltage_readings = [[], [], []]  # List for each phase
                for _ in range(samples):
                    for p in range(1, 4):
                        volt = float(self.agx.query(f'MEAS:VOLT{p}?'))
                        voltage_readings[p-1].append(volt)
                    if samples > 1:
                        time.sleep(0.1)
                
                # Calculate average for each phase
                avg_voltages = [
                    statistics.mean(phase_readings) 
                    for phase_readings in voltage_readings
                ]
                
                if samples > 1:
                    # Calculate standard deviation for each phase
                    std_dev = [
                        statistics.stdev(phase_readings) if len(phase_readings) > 1 else 0
                        for phase_readings in voltage_readings
                    ]
                    return avg_voltages, std_dev
                return avg_voltages, [0, 0, 0]
            else:
                # Measure specific phase
                readings = []
                for _ in range(samples):
                    volt = float(self.agx.query(f'MEAS:VOLT{phase}?'))
                    readings.append(volt)
                    if samples > 1:
                        time.sleep(0.1)
                
                avg = statistics.mean(readings)
                std = statistics.stdev(readings) if len(readings) > 1 else 0
                return avg, std
        except Exception as e:
            print(f"Error measuring voltage: {e}")
            return None

    def enable_output(self, enable: bool = True):
        """Enable or disable AGX output"""
        try:
            if enable:
                self.agx.write('OUTP ON')
                print("Output enabled")
            else:
                self.agx.write('OUTP OFF')
                print("Output disabled")
            return True
        except Exception as e:
            print(f"Error {'enabling' if enable else 'disabling'} output: {e}")
            return False

    def run_test_from_config(self, config: Dict, samples: int = 3):
        """Run test based on configuration dictionary"""
        if not self.connected:
            print("Not connected to AGX")
            return

        # Extract test parameters
        mode = config.get('mode', 'AC')
        phase_config = config.get('phase_config', '3-PHASE')
        frequency = float(config.get('frequency', 50))
        voltage = float(config.get('voltage', 0))
        test_point = config.get('test_point', 'Unknown')
        
        print(f"\n=== Running Test: {test_point} ===")
        print(f"Mode: {mode}")
        print(f"Phase Config: {phase_config}")
        print(f"Frequency: {frequency}Hz")
        print(f"Voltage: {voltage}V")
        
        # Configure AGX
        if not self.configure_mode(mode):
            return
        if not self.setup_phase_config(phase_config, frequency):
            return
            
        # Create results directory if it doesn't exist
        os.makedirs('test_results', exist_ok=True)
        
        # Create CSV file for results
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"test_results/voltage_test_{test_point}_{timestamp}.csv"
        
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([
                'Timestamp',
                'Test Point',
                'Mode',
                'Phase Config',
                'Frequency (Hz)',
                'Set Voltage (V)',
                'Phase 1 Voltage (V)',
                'Phase 1 Std Dev',
                'Phase 2 Voltage (V)',
                'Phase 2 Std Dev',
                'Phase 3 Voltage (V)',
                'Phase 3 Std Dev',
                'Max Deviation (%)'
            ])
            
            # Enable output
            if not self.enable_output(True):
                return

            try:
                # Set voltage
                if self.set_voltage(voltage):
                    # Wait for voltage to stabilize
                    print("Waiting 2 seconds for stabilization...")
                    time.sleep(2)
                    
                    # Measure voltage
                    measured, std_devs = self.measure_voltage(samples=samples)
                    if measured:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        # Calculate maximum deviation from set point as percentage
                        deviations = [abs(v - voltage) / voltage * 100 for v in measured]
                        max_deviation = max(deviations)
                        
                        print("\nMeasurements:")
                        for phase in range(3):
                            print(f"Phase {phase+1}: {measured[phase]:.3f}V ± {std_devs[phase]:.3f}V")
                        print(f"Maximum deviation: {max_deviation:.2f}%")
                        
                        # Write results to CSV
                        writer.writerow([
                            timestamp,
                            test_point,
                            mode,
                            phase_config,
                            frequency,
                            voltage,
                            f"{measured[0]:.3f}",
                            f"{std_devs[0]:.3f}",
                            f"{measured[1]:.3f}",
                            f"{std_devs[1]:.3f}",
                            f"{measured[2]:.3f}",
                            f"{std_devs[2]:.3f}",
                            f"{max_deviation:.2f}"
                        ])
                    else:
                        print("Error measuring voltage")
            finally:
                # Always disable output when done
                self.enable_output(False)
        
        print(f"\nTest complete. Results saved to {filename}")

    def run_tests_from_csv(self, csv_file: str, samples: int = 3):
        """Run multiple tests from a CSV configuration file"""
        print(f"\n=== Running Tests from {csv_file} ===")
        
        try:
            with open(csv_file, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    self.run_test_from_config(row, samples)
        except Exception as e:
            print(f"Error reading CSV file: {e}")

    def disconnect(self):
        """Safely disconnect from AGX"""
        if self.connected:
            try:
                # Turn off output
                self.enable_output(False)
                # Close connection
                self.agx.close()
                print("Disconnected from AGX")
            except Exception as e:
                print(f"Error disconnecting: {e}")
        if self.rm:
            self.rm.close()

def main():
    parser = argparse.ArgumentParser(description='AGX Voltage Test')
    parser.add_argument('--gpib', type=int, default=1, help='GPIB address (default: 1)')
    parser.add_argument('--config', type=str, help='CSV configuration file')
    parser.add_argument('--samples', type=int, default=3,
                       help='Number of samples per measurement (default: 3)')
    
    args = parser.parse_args()
    
    agx = AGXVoltageTest()
    
    # Connect to AGX
    if not agx.connect_to_agx(args.gpib):
        print("Failed to connect to AGX")
        return
    
    try:
        if args.config:
            # Run tests from CSV configuration
            agx.run_tests_from_csv(args.config, args.samples)
        else:
            print("Please provide a CSV configuration file using --config")
    except KeyboardInterrupt:
        print("\nTests interrupted by user")
    finally:
        # Always disconnect properly
        agx.disconnect()

if __name__ == "__main__":
    main()

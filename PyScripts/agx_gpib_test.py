import pyvisa
import time
import sys
from datetime import datetime

class GPIBError(Exception):
    """Custom exception for GPIB communication errors"""
    pass

class AGXGPIBTester:
    def __init__(self):
        try:
            self.rm = pyvisa.ResourceManager()
            resources = self.rm.list_resources()
            gpib_devices = [res for res in resources if 'GPIB' in res]
            
            if not gpib_devices:
                raise GPIBError("No GPIB devices found")
                
            # Print available GPIB devices
            print(f"Available GPIB devices: {gpib_devices}")
            
            self.instrument = self.rm.open_resource(gpib_devices[0])
            self.instrument.timeout = 20000  # Increased timeout for longer operations
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            
            # Initial delay for device to be ready
            print("Waiting for device to initialize...")
            time.sleep(10)
            
            # Clear device
            self.write_command("*CLS")
            time.sleep(2)
            
            # Get device ID
            idn = self.query_command("*IDN?")
            if not idn:
                raise GPIBError("Failed to get device identification")
            print(f"Connected to: {idn}")
            
            print("Basic communication verified")
            
        except Exception as e:
            print(f"Initialization error: {str(e)}")
            if hasattr(self, 'instrument'):
                self.instrument.close()
            if hasattr(self, 'rm'):
                self.rm.close()
            sys.exit(1)

    def write_command(self, cmd, retries=3):
        """Send command with retry logic"""
        for attempt in range(retries):
            try:
                print(f"Sending command: {cmd}")
                self.instrument.write(cmd)
                time.sleep(1)  # Basic delay after command
                return True
            except Exception as e:
                print(f"Error sending command '{cmd}' (attempt {attempt + 1}/{retries}): {str(e)}")
                if attempt < retries - 1:
                    time.sleep(2)  # Wait before retry
                    continue
                return False
        return False

    def query_command(self, cmd, retries=3):
        """Send query with retry logic"""
        for attempt in range(retries):
            try:
                print(f"Sending query: {cmd}")
                response = self.instrument.query(cmd)
                time.sleep(1)  # Basic delay after query
                return response.strip()
            except Exception as e:
                print(f"Error querying '{cmd}' (attempt {attempt + 1}/{retries}): {str(e)}")
                if attempt < retries - 1:
                    time.sleep(2)  # Wait before retry
                    continue
                return None
        return None

    def setup_ac_mode(self):
        """Configure for AC output"""
        commands = [
            "*RST",                 # Reset device
            "*CLS",                 # Clear status
            ":OUTP OFF",           # Safety: ensure output is off
            ":VOLT:MODE AC",       # Set AC mode
            ":FORM 3",             # Three phase mode
            ":VOLT:ALC ON",        # Automatic level control
            ":CURR:LIM 10",        # Current limit 10A
            ":WAVEFORM 1",         # Sine wave
            ":SENS:PATH 0",        # Internal sensing path
            ":FREQ 50",            # 50Hz
            ":VOLT:AC:LIM:MIN 0",
            ":VOLT:AC:LIM:MAX 300",
            ":COUPL DIRECT",
            ":PHAS2 120",          # Phase angles
            ":PHAS3 240",
            ":RANG 0",             # Set to low range for better accuracy at lower voltages
            ":RAMP 0",
            ":VOLT:RANG LOW"       # Explicitly set voltage range to LOW
        ]
        
        print("\nSetting up AC mode...")
        for cmd in commands:
            if not self.write_command(cmd):
                print(f"Failed to execute command: {cmd}")
                return False
            time.sleep(1)  # Delay between commands
        return True

    def setup_dc_mode(self):
        """Configure for DC output"""
        commands = [
            "*RST",                 # Reset device
            "*CLS",                 # Clear status
            ":OUTP OFF",           # Safety: ensure output is off
            ":VOLT:MODE DC",       # Set DC mode
            ":FORM 3",             # Three phase mode
            ":VOLT:ALC ON",        # Automatic level control
            ":CURR:LIM 10",        # Current limit 10A
            ":WAVEFORM 1",         # DC waveform
            ":SENS:PATH 0",        # Internal sensing path
            ":VOLT:DC:LIM:MIN 0",
            ":VOLT:DC:LIM:MAX 425",
            ":COUPL DIRECT",
            ":RANG 0",             # Set to low range for better accuracy at lower voltages
            ":RAMP 0",
            ":VOLT:RANG LOW"       # Explicitly set voltage range to LOW
        ]
        
        print("\nSetting up DC mode...")
        for cmd in commands:
            if not self.write_command(cmd):
                print(f"Failed to execute command: {cmd}")
                return False
            time.sleep(1)  # Delay between commands
        return True

    def set_voltage(self, voltage, mode='AC'):
        """Set voltage and verify"""
        print(f"\nSetting {mode} voltage to {voltage}V")
        
        try:
            # Set appropriate voltage range
            if voltage <= 40:
                self.write_command(":RANG 0")
                self.write_command(":VOLT:RANG LOW")
            else:
                self.write_command(":RANG 1")
                self.write_command(":VOLT:RANG HIGH")
            time.sleep(2)  # Wait for range change to settle
            
            # Set voltage command based on mode
            cmd = f":VOLT {voltage}"
            if not self.write_command(cmd):
                return False
                
            # Enable output
            if not self.write_command(":OUTP ON"):
                return False
                
            return True
            
        except Exception as e:
            print(f"Error setting voltage: {str(e)}")
            return False

    def measure_voltage(self):
        """Measure voltage for all phases"""
        try:
            # Measure actual output
            cmd = ":MEAS:VOLT?"
            response = self.query_command(cmd)
            if response:
                try:
                    # Split response by commas if present, otherwise by spaces
                    if ',' in response:
                        parts = response.split(',')
                    else:
                        parts = response.split()
                    
                    values = []
                    for part in parts:
                        # Clean each part and convert to float
                        clean_part = ''.join(c for c in part if c.isdigit() or c in '.-')
                        try:
                            value = float(clean_part)
                            values.append(value)
                        except ValueError:
                            continue
                    
                    if len(values) >= 3:
                        return values[:3]  # Return first 3 values
                    
                except ValueError:
                    print(f"Invalid voltage measurement: {response}")
            return None
        except Exception as e:
            print(f"Error measuring voltage: {str(e)}")
            return None

    def verify_voltage(self, target_voltage, tolerance=0.1):
        """Verify voltage is at target value"""
        for _ in range(3):  # Try up to 3 times
            measured = self.measure_voltage()
            if measured:
                # Check if all phases are within tolerance
                if all(abs(v - target_voltage) <= target_voltage * tolerance for v in measured):
                    return True
            time.sleep(5)  # Wait before retry
        return False

    def run_voltage_test(self, voltage, mode='AC'):
        """Run a complete voltage test sequence"""
        print(f"\nStarting {mode} voltage test at {voltage}V")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n{timestamp}")
        
        # Set voltage
        if not self.set_voltage(voltage, mode):
            print("Failed to set voltage")
            return None
            
        # Take readings during 15 second stabilization period
        print("Starting 15 second stabilization period with readings:")
        readings = []
        for i in range(5):  # Take 5 readings over 15 seconds
            time.sleep(3)  # Wait 3 seconds between readings
            values = self.measure_voltage()
            if values:
                readings.append(values)
                print(f"Reading {i+1} at {i*3+3}s:")
                for phase, v in enumerate(values, 1):
                    print(f"  Phase {phase}: {v:.3f}V")
        
        if not readings:
            print("Failed to get valid measurements")
            return None
            
        # Use only the final reading for statistics
        if readings:
            final_reading = readings[-1]
            phase_stats = []
            for phase in range(3):
                measured = final_reading[phase]
                deviation = abs(measured - voltage)
                phase_stats.append((measured, deviation))
            
            # Display statistics
            print(f"\nFinal Statistics:")
            for phase, (measured, dev) in enumerate(phase_stats, 1):
                print(f"Phase {phase}:")
                print(f"  Measured voltage: {measured:.3f}V")
                print(f"  Deviation: {dev:.3f}V")
                print(f"  Target voltage: {voltage:.1f}V")
        
        # Set voltage back to 0
        print("Ramping down voltage...")
        self.set_voltage(0, mode)
        time.sleep(5)  # Longer wait for voltage to return to 0
        
        # Return both measurements and statistics for logging
        return (readings[-1] if readings else None, phase_stats if readings else None)

    def run_test_sequence(self):
        """Run complete test sequence"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        results_file = f"voltage_test_results_{timestamp}.csv"
        
        with open(results_file, 'w') as f:
            # Write header with statistics columns
            f.write("Timestamp,Mode,Target_Voltage,")
            for phase in range(1, 4):
                f.write(f"Phase{phase}_Value,Phase{phase}_Avg,Phase{phase}_MaxDev,")
            f.write("\n")
            
            # AC Tests
            print("\n=== Starting AC Test Sequence ===")
            if self.setup_ac_mode():
                # Wait for mode to stabilize
                print("Waiting for AC mode to stabilize...")
                time.sleep(20)  # Longer initial stabilization
                
                for voltage in [10, 25, 50, 75, 100, 115]:
                    result = self.run_voltage_test(voltage, 'AC')
                    if result:
                        measurements, stats = result
                        if measurements and stats:
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            line = f"{timestamp},AC,{voltage},"
                            # Add measurements and statistics for each phase
                            for phase in range(3):
                                avg, dev = stats[phase]
                                line += f"{measurements[phase]:.3f},{avg:.3f},{dev:.3f},"
                            f.write(line.rstrip(',') + '\n')
                    # Longer delay between tests
                    print("Waiting between tests...")
                    time.sleep(10)
                
                # Ensure output is off and voltage is 0 before mode change
                self.write_command(":VOLT 0")
                time.sleep(5)
                self.write_command(":OUTP OFF")
                time.sleep(10)  # Longer delay before mode change
                    
            # DC Tests
            print("\n=== Starting DC Test Sequence ===")
            if self.setup_dc_mode():
                # Wait for mode to stabilize
                print("Waiting for DC mode to stabilize...")
                time.sleep(20)  # Longer initial stabilization
                
                for voltage in [10, 25, 50, 75, 100]:
                    result = self.run_voltage_test(voltage, 'DC')
                    if result:
                        measurements, stats = result
                        if measurements and stats:
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            line = f"{timestamp},DC,{voltage},"
                            # Add measurements and statistics for each phase
                            for phase in range(3):
                                avg, dev = stats[phase]
                                line += f"{measurements[phase]:.3f},{avg:.3f},{dev:.3f},"
                            f.write(line.rstrip(',') + '\n')
                    # Longer delay between tests
                    print("Waiting between tests...")
                    time.sleep(10)

    def shutdown(self):
        """Safe shutdown sequence"""
        try:
            print("\nPerforming safe shutdown...")
            self.write_command(":VOLT 0")
            time.sleep(1)
            self.write_command(":OUTP OFF")
            time.sleep(1)
            self.write_command("*RST")
            time.sleep(1)
            self.instrument.close()
            self.rm.close()
            print("Shutdown complete")
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")

def main():
    tester = None
    try:
        tester = AGXGPIBTester()
        
        print("\nAGX GPIB Voltage Test Sequence")
        print("=" * 50)
        print("This script will run through voltage test sequences.")
        print("Please ensure all safety measures are in place.")
        print("\nPress Enter to begin...")
        input()
        
        tester.run_test_sequence()
        
    except KeyboardInterrupt:
        print("\nTest sequence interrupted by user")
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        if tester:
            tester.shutdown()

if __name__ == "__main__":
    main()

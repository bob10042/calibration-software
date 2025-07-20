import os
import sys
import time
import subprocess
import pyvisa
import ipaddress
import platform
import requests
from typing import Optional
import socket
import netifaces

class AGXSetupUtility:
    def __init__(self):
        self.rm = None
        self.instrument = None
        self.connection_type = None
        
    def print_header(self, text: str):
        """Print formatted header text"""
        print("\n" + "="*80)
        print(f" {text} ")
        print("="*80 + "\n")

    def get_user_input(self, prompt: str, valid_options: list = None) -> str:
        """Get validated user input"""
        while True:
            response = input(f"{prompt}: ").strip().lower()
            if valid_options is None or response in valid_options:
                return response
            print(f"Please enter one of: {', '.join(valid_options)}")

    def check_admin_rights(self) -> bool:
        """Check if script has admin rights"""
        try:
            return os.getuid() == 0
        except AttributeError:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin() != 0

    def install_required_packages(self):
        """Install required Python packages"""
        self.print_header("Installing Required Packages")
        packages = ['pyvisa', 'pyvisa-py', 'requests', 'netifaces']
        
        for package in packages:
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                print(f"Successfully installed {package}")
            except subprocess.CalledProcessError:
                print(f"Failed to install {package}. Please install it manually.")
                sys.exit(1)

    def configure_network_adapter(self, ip: str, subnet: str, connection_type: str):
        """Configure network adapter settings"""
        self.print_header(f"Configuring Network Adapter for {connection_type}")
        
        if not self.check_admin_rights():
            print("This script requires administrator privileges to configure network settings.")
            print("Please run it as administrator/root.")
            sys.exit(1)

        # Get list of network adapters
        adapters = netifaces.interfaces()
        
        print("\nAvailable network adapters:")
        for idx, adapter in enumerate(adapters):
            print(f"{idx + 1}. {adapter}")
        
        adapter_idx = int(self.get_user_input("\nSelect the adapter number to configure")) - 1
        selected_adapter = adapters[adapter_idx]

        if platform.system() == 'Windows':
            # Windows network configuration
            commands = [
                f'netsh interface ip set address name="{selected_adapter}" static {ip} {subnet}',
            ]
            
            for cmd in commands:
                try:
                    subprocess.run(cmd, shell=True, check=True)
                    print(f"Successfully executed: {cmd}")
                except subprocess.CalledProcessError as e:
                    print(f"Error configuring network: {e}")
                    return False
        else:
            # Linux/Unix network configuration
            try:
                subprocess.run(f'sudo ifconfig {selected_adapter} {ip} netmask {subnet}', shell=True, check=True)
                print("Network configuration successful")
            except subprocess.CalledProcessError as e:
                print(f"Error configuring network: {e}")
                return False

        return True

    def test_connection(self, ip: str) -> bool:
        """Test network connection using ping"""
        self.print_header(f"Testing Connection to {ip}")
        
        ping_param = '-n' if platform.system().lower() == 'windows' else '-c'
        command = ['ping', ping_param, '4', ip]
        
        try:
            subprocess.run(command, check=True)
            print(f"Successfully pinged {ip}")
            return True
        except subprocess.CalledProcessError:
            print(f"Failed to ping {ip}")
            return False

    def setup_usb_lan_emulation(self):
        """Setup USB-LAN emulation mode"""
        self.print_header("USB-LAN Emulation Mode Setup")
        
        print("Steps to enable LAN emulation mode on AGX:")
        print("1. Press SYST (System Menu)")
        print("2. Navigate to INTERFACE SETUP")
        print("3. Select USB Device Interface")
        print("4. Enable LAN Emulation")
        
        input("\nPress Enter once you've completed these steps...")
        
        # Configure computer's network adapter
        if not self.configure_network_adapter('192.168.123.2', '255.255.255.0', 'USB-LAN'):
            return False
        
        # Test connection
        if not self.test_connection('192.168.123.1'):
            return False
        
        # Test web interface
        try:
            response = requests.get('http://192.168.123.1', timeout=5)
            if response.status_code == 200:
                print("Successfully connected to AGX web interface")
                return True
        except requests.RequestException:
            print("Failed to connect to AGX web interface")
            return False

    def setup_direct_lan(self):
        """Setup direct LAN connection"""
        self.print_header("Direct LAN Connection Setup")
        
        print("Steps to configure LAN settings on AGX:")
        print("1. Press SYST (System Menu)")
        print("2. Navigate to INTERFACE SETUP")
        print("3. Select LAN Interface")
        print("4. Set IP Mode to Static")
        print("5. Configure IP: 192.168.1.100")
        print("6. Configure Subnet: 255.255.255.0")
        
        input("\nPress Enter once you've completed these steps...")
        
        # Configure computer's network adapter
        if not self.configure_network_adapter('192.168.1.2', '255.255.255.0', 'Direct LAN'):
            return False
        
        # Test connection
        if not self.test_connection('192.168.1.100'):
            return False
        
        # Test web interface
        try:
            response = requests.get('http://192.168.1.100', timeout=5)
            if response.status_code == 200:
                print("Successfully connected to AGX web interface")
                return True
        except requests.RequestException:
            print("Failed to connect to AGX web interface")
            return False

    def setup_gpib(self):
        """Setup GPIB connection"""
        self.print_header("GPIB Connection Setup")
        
        # Check for NI-VISA installation
        try:
            self.rm = pyvisa.ResourceManager()
        except:
            print("Error: NI-VISA not found. Please install NI-VISA from National Instruments website.")
            return False

        print("Steps to configure GPIB on AGX:")
        print("1. Press SYST → INTERFACE SETUP → GPIB Interface")
        print("2. Set GPIB Status to Enabled")
        print("3. Set GPIB Address (default: 1)")
        
        gpib_address = self.get_user_input("Enter the GPIB address set on the AGX")
        
        try:
            # Try to open GPIB connection
            self.instrument = self.rm.open_resource(f'GPIB0::{gpib_address}::INSTR')
            
            # Test communication with *IDN? query
            response = self.instrument.query('*IDN?')
            print(f"Successfully connected to: {response}")
            return True
            
        except pyvisa.Error as e:
            print(f"Error connecting to GPIB device: {e}")
            return False

    def run_scpi_tests(self):
        """Run basic SCPI commands to test communication"""
        self.print_header("Running SCPI Tests")
        
        if not self.instrument:
            print("No instrument connection established")
            return False
        
        try:
            # Basic device identification
            idn = self.instrument.query('*IDN?')
            print(f"Device identification: {idn}")
            
            # Check communication interfaces
            print("\nChecking communication interfaces...")
            
            # Check USB status
            try:
                usb_status = self.instrument.query('SYSTem:COMMunicate:USB:VIRTualport:ENABle?')
                print(f"USB Virtual Port Status: {usb_status}")
            except:
                print("Could not query USB status")
            
            # Check LAN status
            try:
                lan_status = self.instrument.query('SYSTem:COMMunicate:LAN:STATus?')
                print(f"LAN Status: {lan_status}")
            except:
                print("Could not query LAN status")
            
            # Check GPIB status
            try:
                gpib_status = self.instrument.query('SYSTem:COMMunicate:GPIB:ENABle?')
                print(f"GPIB Status: {gpib_status}")
            except:
                print("Could not query GPIB status")
            
            return True
            
        except pyvisa.Error as e:
            print(f"Error during SCPI tests: {e}")
            return False

    def main(self):
        """Main setup procedure"""
        self.print_header("AGX Setup Utility")
        
        # Install required packages
        self.install_required_packages()
        
        # Choose connection type
        print("\nAvailable connection types:")
        print("1. USB-LAN Emulation Mode")
        print("2. Direct LAN Connection")
        print("3. GPIB Connection")
        
        choice = self.get_user_input("Select connection type (1-3)", ['1', '2', '3'])
        
        success = False
        if choice == '1':
            success = self.setup_usb_lan_emulation()
        elif choice == '2':
            success = self.setup_direct_lan()
        elif choice == '3':
            success = self.setup_gpib()
        
        if success:
            print("\nConnection established successfully!")
            if choice == '3':  # Only run SCPI tests for GPIB
                self.run_scpi_tests()
        else:
            print("\nFailed to establish connection. Please check the following:")
            print("1. AGX is powered on and properly connected")
            print("2. Correct drivers are installed")
            print("3. Network/GPIB settings are correct")
            print("4. No firewall blocking the connection")

if __name__ == "__main__":
    setup = AGXSetupUtility()
    setup.main()

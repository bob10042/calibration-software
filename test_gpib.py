import pyvisa
import time

def test_gpib_connection():
    try:
        # Create a resource manager
        rm = pyvisa.ResourceManager()
        
        # List all available resources
        resources = rm.list_resources()
        print("\nAvailable Resources:")
        for resource in resources:
            print(f"- {resource}")
            
        # If no resources found
        if not resources:
            print("No VISA resources found. Please check if device is connected properly.")
            return
            
        # Try to connect to first GPIB device found
        gpib_devices = [res for res in resources if 'GPIB' in res]
        if not gpib_devices:
            print("No GPIB devices found.")
            return
            
        # Connect to first GPIB device
        print(f"\nAttempting to connect to: {gpib_devices[0]}")
        instrument = rm.open_resource(gpib_devices[0])
        
        # Set timeout to 5 seconds
        instrument.timeout = 5000
        
        # Query device identification
        print("Sending *IDN? query...")
        idn = instrument.query("*IDN?")
        print(f"Device responded: {idn}")
        
        # Close connection
        instrument.close()
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        # Clean up
        if 'instrument' in locals():
            instrument.close()
        if 'rm' in locals():
            rm.close()

if __name__ == "__main__":
    print("Starting GPIB Connection Test...")
    test_gpib_connection()

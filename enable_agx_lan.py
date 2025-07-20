import serial
import time

def send_command(ser, cmd):
    print(f"Sending: {cmd}")
    ser.write(f"{cmd}\n".encode())
    time.sleep(0.5)  # Give device time to process
    response = ser.read_all().decode().strip()
    print(f"Response: {response}")
    return response

try:
    # Open serial connection
    ser = serial.Serial('COM3', baudrate=9600, timeout=1)
    print("Connected to COM3")
    
    # Enable virtual COM port
    send_command(ser, "SYSTem:COMMunicate:USB:VIRTualport:ENABle 1")
    
    # Enable USB-LAN interface
    send_command(ser, "SYSTem:COMMunicate:USB:LAN:ENABle 1")
    
    # Query LAN status
    send_command(ser, "SYSTem:COMMunicate:LAN:STATus?")
    
    ser.close()
    print("Done")

except Exception as e:
    print(f"Error: {str(e)}")

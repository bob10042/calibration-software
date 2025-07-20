import serial
import time

def send_command(ser, cmd):
    print(f"Sending: {cmd}")
    ser.write(f"{cmd}\n".encode())
    time.sleep(0.5)
    response = ser.read_all().decode().strip()
    print(f"Response: {response}")
    return response

try:
    ser = serial.Serial('COM3', baudrate=9600, timeout=1)
    print("Connected to COM3")
    
    # Set subnet mask to match our network
    send_command(ser, "SYSTem:COMMunicate:LAN:SMASk 255.255.252.0")
    
    # Set gateway
    send_command(ser, "SYSTem:COMMunicate:LAN:GATeway 192.168.40.254")
    
    # Query LAN status to verify changes
    send_command(ser, "SYSTem:COMMunicate:LAN:STATus?")
    
    ser.close()
    print("Done")

except Exception as e:
    print(f"Error: {str(e)}")

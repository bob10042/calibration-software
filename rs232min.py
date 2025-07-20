import time
import serial  # pip install pyserial

def init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False):
    """
    Example of communicating with the APSM2000 (M2000 Series) via RS232 (or USB–RS232 adapter).
    1) Opens the serial port (assigned by OS).
    2) Sends *CLS to clear interface.
    3) (Optionally) LOCKOUT to prevent front-panel changes.
    4) Requests *IDN?.
    5) Prints and parses the returned ID string.
    6) Closes the serial port.
    """
    ser = serial.Serial()
    ser.port = port_name
    ser.baudrate = baud_rate

    # 8 data bits, no parity, 1 stop bit
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0  # seconds read timeout

    # If APSM2000 is configured for hardware flow control:
    ser.rtscts = True
    # Assert DTR so APSM2000 recognizes a controller is present:
    ser.dtr = True

    try:
        ser.open()
        print(f"Opened {port_name} at {baud_rate} baud.")

        # Clear interface
        ser.write(b"*CLS\n")
        time.sleep(0.1)

        # (Optional) Lock out front panel
        if lock_front_panel:
            ser.write(b"LOCKOUT\n")
            time.sleep(0.1)

        # Request ID
        ser.write(b"*IDN?\n")
        response = ser.readline().decode('ascii', errors='replace').strip()
        print("Raw *IDN? response:", response)

        # Parse the comma-separated fields
        fields = response.split(',')
        print("Parsed *IDN? fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # Example of reading 3-channel voltages:
        # ser.write(b"READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3\n")
        # volt_resp = ser.readline().decode('ascii').strip()
        # print("3-Ch Voltage response:", volt_resp)

        # Revert to local if you want to unlock panel:
        # ser.write(b"LOCAL\n")

    except serial.SerialException as e:
        print("Serial error:", e)
    except Exception as e:
        print("General exception:", e)
    finally:
        if ser.is_open:
            ser.close()
            print(f"Closed {port_name}")

if __name__ == "__main__":
    init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False)

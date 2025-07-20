import time
import serial  # Requires: pip install pyserial

def init_and_read_idn_rs232(
    port_name="COM4", 
    baud_rate=115200, 
    lock_front_panel=False
):
    """
    Example of communicating with the M2000 (APSM2000) via a USB-to-RS232 adapter.
    1) Opens the serial port (assigned by OS).
    2) Sends *CLS to clear interface.
    3) (Optionally) LOCKOUT to prevent front-panel changes.
    4) Requests *IDN?.
    5) Prints and parses the returned ID string.
    6) Closes the serial port.
    
    Adjust 'port_name' to match your system's COM port.
    """
    ser = serial.Serial()
    ser.port = port_name
    ser.baudrate = baud_rate
    # 8N1, no parity, 1 stop bit
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0  # seconds read timeout

    # If the M2000 is configured for hardware flow control (RTS/CTS):
    ser.rtscts = True
    # DTR must be asserted so the M2000 sees an active controller:
    ser.dtr = True

    try:
        ser.open()
        print(f"Opened {port_name} at {baud_rate} baud.")

        # 1) Clear interface / flush errors:
        ser.write(b"*CLS\n")
        time.sleep(0.1)

        # 2) (Optional) Lock out the front panel
        if lock_front_panel:
            ser.write(b"LOCKOUT\n")
            time.sleep(0.1)

        # 3) Request ID
        ser.write(b"*IDN?\n")
        response = ser.readline().decode('ascii', errors='replace').strip()
        print("Raw *IDN? response:", response)

        # Parse the comma-separated fields
        fields = response.split(',')
        # Typically up to 6 fields: manufacturer, model(+options), serial#, fw_major, fw_minor, fw_build
        print("Parsed fields:")
        for i, f in enumerate(fields, start=1):
            print(f"  Field{i} = {f}")

        # If you want to revert to local from software (unlock front panel):
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
    # Adjust the COM port name and baud rate to match your system & M2000 config
    init_and_read_idn_rs232(port_name="COM4", baud_rate=115200, lock_front_panel=False)

import serial
import time

def read_voltages_via_rs232(port_name="COM1", baud_rate=115200):
    """
    Example of opening an RS232 port via pyserial,
    sending a command to read all channel voltages,
    and printing the response.
    """
    # Configure the serial port
    ser = serial.Serial()
    ser.port = port_name       # e.g., "COM1", "COM4", "/dev/ttyUSB0", etc.
    ser.baudrate = baud_rate   # Must match your M2000 setting
    ser.bytesize = serial.EIGHTBITS
    ser.parity   = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.timeout  = 1.0         # seconds read timeout
    # Enable RTS/CTS hardware flow control:
    ser.rtscts   = True
    # DTR must be enabled so M2000 knows a controller is present:
    ser.dtr      = True

    try:
        ser.open()
        # The M2000 goes into remote mode once it receives any valid command.
        # Clear interface and errors:
        ser.write(b"*CLS\n")
        time.sleep(0.2)

        # Example command: read the voltage on channels 1..4.
        # Using the syntax:  READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3, VOLTS:CH4
        command = "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3, VOLTS:CH4\n"
        ser.write(command.encode("ascii"))

        # Read back one line of response (terminated by CR,LF typically)
        response_line = ser.readline().decode("ascii").strip()

        print("Raw response:", response_line)

        # The M2000 typically returns comma-separated NR3 fields, e.g.:
        # +1.20000E+01,+2.34567E+01,+3.21000E+01,+4.54321E+01
        # We can split on commas:
        if response_line:
            voltage_strings = response_line.split(',')
            voltages = [float(v) for v in voltage_strings]
            print("Parsed channel voltages:", voltages)
        else:
            print("No response or empty response received.")

    except serial.SerialException as e:
        print("Serial error:", e)
    except Exception as e:
        print("General error:", e)
    finally:
        if ser.is_open:
            ser.close()

if __name__ == "__main__":
    read_voltages_via_rs232(port_name="COM3", baud_rate=115200)

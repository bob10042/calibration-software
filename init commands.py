        # Simplified initialization commands
        send_command(ser, "*IDN?")  # Query device ID
        idn_response = read_response(ser)
        if idn_response:
            logging.info(f"[INFO] IDN Response: {idn_response}")
            print(f"[INFO] IDN Response: {idn_response}")
        else:
            logging.warning("[WARNING] No response for *IDN? command.")
            return

        # Test Single VPA Mode
        print("Testing Single VPA Mode...")
        send_command(ser, "MODE,0")
        mode_response = read_response(ser)
        if mode_response:
            logging.info(f"[INFO] MODE Response: {mode_response}")
            print(f"[INFO] MODE Response: {mode_response}")
        else:
            logging.warning("[WARNING] No response for MODE command.")
            return

        # Test basic channel reads
        print("Testing single channel read...")
        send_command(ser, "READ? VOLTS:CH1")
        ch1_response = read_response(ser)
        if ch1_response:
            logging.info(f"[INFO] CH1 Voltage: {ch1_response}")
            print(f"[INFO] CH1 Voltage: {ch1_response}")
        else:
            logging.warning("[WARNING] No response for CH1 voltage read.")
            return

        # Test reading all channels
        print("Testing all channels...")
        send_command(ser, "READ? VOLTS:CH1, VOLTS:CH2, VOLTS:CH3")
        all_channels_response = read_response(ser)
        if all_channels_response:
            logging.info(f"[INFO] All Channels Voltage: {all_channels_response}")
            print(f"[INFO] All Channels Voltage: {all_channels_response}")
        else:
            logging.warning("[WARNING] No response for all channels voltage read.")

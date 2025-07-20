# SCPI Commands for Setting AC/DC Modes and Current

This document provides SCPI commands to configure and control the AGX Series programmable power source, focusing on setting different modes and managing AC/DC outputs.

---

## 1. **Initialize the Power Source**
Before starting any configuration, initialize the system:

### a. Reset and Clear Errors:
```plaintext
*RST  // Reset the power source to default settings.
*CLS  // Clear any existing errors or events.
```

### b. Verify Initialization:
```plaintext
SYSTem:ERRor?  // Query the system for errors to ensure it is initialized correctly.
```

---

## 2. **Ensure Voltage Display on Machine**
To ensure that the voltage is displayed on the machine:

### a. Enable Voltage Display:
```plaintext
DISPlay:VOLTage ON  // Turn on voltage display.
```

### b. Verify Display:
```plaintext
DISPlay:STATe?  // Check the display state to confirm voltage is being shown.
```

---

## 3. **Setting AC Mode**
Configure the system to output AC power.

### a. Set AC Voltage:
```plaintext
SOURce:VOLTage:AC <voltage_value>  // Example: 230
```

### b. Set AC Frequency:
```plaintext
SOURce:FREQuency <frequency_value>  // Example: 50
```

### c. Enable AC Output:
```plaintext
SOURce:MODE AC
OUTPut:STATe ON
```

---

## 4. **Setting DC Mode**
Configure the system to output DC power.

### a. Set DC Voltage:
```plaintext
SOURce:VOLTage:DC <voltage_value>  // Example: 48
```

### b. Enable DC Output:
```plaintext
SOURce:MODE DC
OUTPut:STATe ON
```

---

## 5. **Setting AC + DC Combined Mode**
Combine AC and DC outputs.

### a. Set AC and DC Components:
```plaintext
SOURce:VOLTage:AC <ac_voltage_value>  // Example: 120
SOURce:VOLTage:DC <dc_voltage_value>  // Example: 10
```

### b. Enable Combined Output:
```plaintext
SOURce:MODE AC+DC
OUTPut:STATe ON
```

---

## 6. **Configuring Current Settings**

### a. Set Current for AC Mode:
```plaintext
SOURce:CURRent:AC <current_value>  // Example: 5
```

### b. Set Current for DC Mode:
```plaintext
SOURce:CURRent:DC <current_value>  // Example: 3
```

### c. Enable Current Limit Protection:
```plaintext
SOURce:CURRent:PROTection ON
SOURce:CURRent:PROTection:LEVel <limit_value>  // Example: 10
```

---

## 7. **Querying Current and Voltage**
Retrieve real-time values for monitoring.

### a. Query Output Voltage:
```plaintext
MEASure:VOLTage:AC?  // For AC voltage
MEASure:VOLTage:DC?  // For DC voltage
```

### b. Query Output Current:
```plaintext
MEASure:CURRent:AC?  // For AC current
MEASure:CURRent:DC?  // For DC current
```

---

## 8. **Error Management with Real-Time Feedback**
Handle and query errors during operation and enable real-time error reporting:

### a. Enable Real-Time Error Feedback:
```plaintext
STATus:QUEStionable:ENABle <mask_value>  // Enable real-time error monitoring.
*ESE 1  // Enable error event reporting to the Event Status Register.
*OPC  // Ensure the system is ready for commands.
```

### b. Check for Errors:
```plaintext
SYSTem:ERRor?  // Query the error queue for any existing errors.
```

### c. Clear Errors:
```plaintext
*CLS  // Clear the error queue.
```

### d. Log Errors (Optional):
```plaintext
SYSTem:ERRor:NEXT?  // Retrieve the next error message from the queue.
```

### e. Query Error Description:
```plaintext
SYSTem:ERRor:DESCription? <error_code>  // Get details for a specific error code.
```

### f. Disable Real-Time Error Feedback (Optional):
```plaintext
STATus:QUEStionable:ENABle 0  // Disable real-time error monitoring.
```

---

## 9. **Terminating the Test**
To safely terminate the test and turn off the output:

### a. Disable the Output:
```plaintext
OUTPut:STATe OFF
```

### b. Reset the System (Optional):
```plaintext
*RST  // Reset to default settings.
```

---

## 10. **General System Commands**

### a. Check System Status:
```plaintext
STATus:QUEStionable:CONDition?
```

### b. Check Errors:
```plaintext
SYSTem:ERRor?
```

---

### Notes:
1. Replace `<value>` placeholders with appropriate numerical values based on your requirements.
2. Ensure safe operating conditions and verify settings before enabling outputs.
3. Commands are case-insensitive but are written in uppercase for clarity.

---

Save this document for quick reference when working with SCPI commands for the AGX Series power source.


using System;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Microsoft.Win32;
using NationalInstruments.Visa;

namespace AGXCalibrationUI.Services
{
    public class PpsDeviceCommunicationService : IDisposable
    {
        private SerialPort? _serialPort;
        private MessageBasedSession? _visaSession;
        private bool _isConnected;
        private const int DefaultBaudRate = 9600;
        private const int DefaultTimeout = 2000;  // 2 seconds
        private const string TerminationCharacter = "\n";

        public bool IsConnected => _isConnected;

        public PpsDeviceCommunicationService()
        {
            _serialPort = new SerialPort
            {
                BaudRate = DefaultBaudRate,
                DataBits = 8,
                Parity = Parity.None,
                StopBits = StopBits.One,
                ReadTimeout = DefaultTimeout,
                WriteTimeout = DefaultTimeout,
                NewLine = TerminationCharacter
            };
        }

        public async Task<List<string>> GetAvailableDevicesAsync()
        {
            var devices = new List<string>();
            
            try
            {
                // Get all available COM ports
                string[] ports = SerialPort.GetPortNames();
                Debug.WriteLine($"Found {ports.Length} COM ports");

                foreach (string port in ports)
                {
                    try
                    {
                        using (var testPort = new SerialPort(port))
                        {
                            testPort.BaudRate = DefaultBaudRate;
                            testPort.DataBits = 8;
                            testPort.Parity = Parity.None;
                            testPort.StopBits = StopBits.One;
                            testPort.ReadTimeout = DefaultTimeout;
                            testPort.WriteTimeout = DefaultTimeout;
                            testPort.NewLine = TerminationCharacter;

                            testPort.Open();
                            if (testPort.IsOpen)
                            {
                                // Try to identify if this is a PPS device
                                testPort.WriteLine("*IDN?");
                                await Task.Delay(100);  // Give device time to respond
                                string? response = null;

                                try
                                {
                                    response = testPort.ReadLine();
                                }
                                catch (TimeoutException)
                                {
                                    Debug.WriteLine($"Timeout reading from {port}");
                                }

                                if (response != null && 
                                    (response.Contains("PPS", StringComparison.OrdinalIgnoreCase) ||
                                     response.Contains("AGX", StringComparison.OrdinalIgnoreCase)))
                                {
                                    Debug.WriteLine($"Found PPS device on {port}: {response}");
                                    devices.Add($"COM:{port}");
                                }
                            }
                            testPort.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error checking {port}: {ex.Message}");
                    }
                }

                // Get all available USB devices using VISA
                try
                {
                    using (var resourceManager = new ResourceManager())
                    {
                        var resources = resourceManager.Find("USB?*");
                        foreach (string resource in resources)
                        {
                            try
                            {
                                using (var session = (MessageBasedSession)resourceManager.Open(resource))
                                {
                                    session.TimeoutMilliseconds = DefaultTimeout;
                                    
                                    // Write the identification query
                                    session.RawIO.Write("*IDN?\n");
                                    string response = session.RawIO.ReadString().Trim();

                                    if (response.Contains("PPS", StringComparison.OrdinalIgnoreCase) ||
                                        response.Contains("AGX", StringComparison.OrdinalIgnoreCase))
                                    {
                                        Debug.WriteLine($"Found PPS USB device: {resource} - {response}");
                                        devices.Add($"USB:{resource}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Error checking USB device {resource}: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error enumerating USB devices: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error enumerating devices: {ex.Message}");
                throw;
            }

            return devices;
        }

        public async Task<bool> ConnectAsync(string deviceAddress)
        {
            try
            {
                if (_isConnected)
                {
                    await DisconnectAsync();
                }

                // Parse device address to determine connection type
                var parts = deviceAddress.Split(new[] { ':' }, 2);
                if (parts.Length != 2)
                {
                    throw new ArgumentException("Invalid device address format. Expected format: TYPE:ADDRESS");
                }

                var type = parts[0].ToUpperInvariant();
                var address = parts[1];

                switch (type)
                {
                    case "COM":
                        return await ConnectSerialAsync(address);
                    case "USB":
                        return await ConnectUSBAsync(address);
                    default:
                        throw new ArgumentException($"Unsupported connection type: {type}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Connection failed: {ex.Message}");
                throw;
            }
        }

        private async Task<bool> ConnectSerialAsync(string portName)
        {
            try
            {
                if (_serialPort == null)
                {
                    _serialPort = new SerialPort
                    {
                        BaudRate = DefaultBaudRate,
                        DataBits = 8,
                        Parity = Parity.None,
                        StopBits = StopBits.One,
                        ReadTimeout = DefaultTimeout,
                        WriteTimeout = DefaultTimeout,
                        NewLine = TerminationCharacter
                    };
                }

                _serialPort.PortName = portName;
                _serialPort.Open();

                if (_serialPort.IsOpen)
                {
                    // Verify it's a PPS device
                    _serialPort.WriteLine("*IDN?");
                    await Task.Delay(100);
                    string response = _serialPort.ReadLine().Trim();

                    if (response.Contains("PPS", StringComparison.OrdinalIgnoreCase) ||
                        response.Contains("AGX", StringComparison.OrdinalIgnoreCase))
                    {
                        _isConnected = true;
                        Debug.WriteLine($"Connected to PPS device via COM: {response}");
                        return true;
                    }
                    else
                    {
                        _serialPort.Close();
                        throw new Exception("Connected device is not a PPS/AGX device");
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                if (_serialPort?.IsOpen == true)
                {
                    _serialPort.Close();
                }
                throw new Exception($"Serial connection failed: {ex.Message}", ex);
            }
        }

        private async Task<bool> ConnectUSBAsync(string resourceName)
        {
            try
            {
                var resourceManager = new ResourceManager();
                _visaSession = (MessageBasedSession)resourceManager.Open(resourceName);
                _visaSession.TimeoutMilliseconds = DefaultTimeout;

                // Write identification query and read response
                _visaSession.RawIO.Write("*IDN?\n");
                string response = _visaSession.RawIO.ReadString().Trim();

                if (response.Contains("PPS", StringComparison.OrdinalIgnoreCase) ||
                    response.Contains("AGX", StringComparison.OrdinalIgnoreCase))
                {
                    _isConnected = true;
                    Debug.WriteLine($"Connected to PPS device via USB: {response}");
                    return true;
                }
                else
                {
                    _visaSession.Dispose();
                    _visaSession = null;
                    throw new Exception("Connected device is not a PPS/AGX device");
                }
            }
            catch (Exception ex)
            {
                if (_visaSession != null)
                {
                    _visaSession.Dispose();
                    _visaSession = null;
                }
                throw new Exception($"USB connection failed: {ex.Message}", ex);
            }
        }

        public async Task DisconnectAsync()
        {
            if (_serialPort?.IsOpen == true)
            {
                try
                {
                    _serialPort.Close();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error during serial disconnect: {ex.Message}");
                }
            }

            if (_visaSession != null)
            {
                try
                {
                    _visaSession.Dispose();
                    _visaSession = null;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error during USB disconnect: {ex.Message}");
                }
            }

            _isConnected = false;
        }

        public async Task WriteAsync(string command)
        {
            if (!_isConnected)
            {
                throw new Exception("Not connected to device");
            }

            try
            {
                Debug.WriteLine($"Writing command: {command}");
                
                if (_serialPort?.IsOpen == true)
                {
                    await Task.Run(() => _serialPort.WriteLine(command));
                }
                else if (_visaSession != null)
                {
                    await Task.Run(() => _visaSession.RawIO.Write(command + "\n"));
                }
                else
                {
                    throw new Exception("No valid connection available");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Write failed: {ex.Message}");
                throw;
            }
        }

        public async Task<string> ReadAsync()
        {
            if (!_isConnected)
            {
                throw new Exception("Not connected to device");
            }

            try
            {
                string response;
                
                if (_serialPort?.IsOpen == true)
                {
                    response = await Task.Run(() => _serialPort.ReadLine());
                }
                else if (_visaSession != null)
                {
                    response = await Task.Run(() => _visaSession.RawIO.ReadString());
                }
                else
                {
                    throw new Exception("No valid connection available");
                }

                Debug.WriteLine($"Read response: {response}");
                return response.Trim();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Read failed: {ex.Message}");
                throw;
            }
        }

        public async Task<string> QueryAsync(string command)
        {
            if (_visaSession != null)
            {
                // For USB devices, write command and read response
                await Task.Run(() => _visaSession.RawIO.Write(command + "\n"));
                await Task.Delay(100); // Give device time to respond
                return await Task.Run(() => _visaSession.RawIO.ReadString().Trim());
            }
            else
            {
                // Fallback to manual Write/Read for serial connections
                await WriteAsync(command);
                await Task.Delay(100);  // Give device time to respond
                return await ReadAsync();
            }
        }

        public void Dispose()
        {
            if (_serialPort != null)
            {
                if (_serialPort.IsOpen)
                {
                    _serialPort.Close();
                }
                _serialPort.Dispose();
                _serialPort = null;
            }

            if (_visaSession != null)
            {
                _visaSession.Dispose();
                _visaSession = null;
            }

            _isConnected = false;
        }
    }
}
